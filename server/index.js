import 'dotenv/config';
import express from 'express';
import { writeFileSync, readFileSync, existsSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import cors from 'cors';
import { fetchWithPuppeteer } from './fetcher.js';
import { parseJournalList, parseJournalPage, parseJournalDetails } from './parser.js';
import { getClassroomClient, createCourse, addTeacher, addStudentToCourse, removeStudentFromCourse, removeTeacherFromCourse, archiveCourse, updateCourse, getAuthUrl, getTokensFromCode, getTeachersFromAdmin, getClassesFromAdmin, getUsersFromOrgUnit, getUserInfo, moveUserToOrgUnit, createUserInOrgUnit, searchUsersInDomain, createUserInOrgPath, moveUserToOrgPath, updateUser, listCourses, getCourseTeachers, getCourseStudents, updateOrgUnit, deleteOrgUnit } from './classroom.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(cors());
app.use(express.json());

const teacherEmailMap = new Map();
const UA_TO_LATIN = { 'а':'a','б':'b','в':'v','г':'h','ґ':'g','д':'d','е':'e','є':'ie','ж':'zh','з':'z','и':'y','і':'i','ї':'i','й':'i','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'shch','ь':'','ю':'iu','я':'ia','ё':'io' };
const CREATED_STUDENTS_FILE = join(dirname(fileURLToPath(import.meta.url)), '..', 'created-students.json');
let createdStudentsStore = [];
try {
  if (existsSync(CREATED_STUDENTS_FILE)) {
    createdStudentsStore = JSON.parse(readFileSync(CREATED_STUDENTS_FILE, 'utf8'));
  }
} catch (_) {}

function transliterateUAtoEN(s) {
  return (s || '').toLowerCase().split('').map(c => {
    const lower = c.toLowerCase();
    if (UA_TO_LATIN[lower]) return UA_TO_LATIN[lower];
    if (/[a-z0-9]/.test(c)) return c;
    return '';
  }).join('');
}

function parseStudentName(fullName) {
  const parts = (fullName || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) return { familyName: parts[0], givenName: parts.slice(1).join(' ') };
  if (parts.length === 1) return { familyName: parts[0], givenName: '' };
  return { familyName: '', givenName: '' };
}

function generateStudentEmail(familyName, givenName, existingEmails = new Set()) {
  const fn = transliterateUAtoEN((familyName || '').replace(/\s+/g, ' ').trim());
  const gn = transliterateUAtoEN((givenName || '').replace(/\s+/g, ' ').trim());
  const base = [fn, gn].filter(Boolean).join('_').replace(/[^a-z0-9_]/g, '') || 'user';
  const domain = process.env.STUDENT_EMAIL_DOMAIN || 'kshg.site';
  let email = `${base}@${domain}`;
  let n = 1;
  while (existingEmails.has(email.toLowerCase())) {
    email = `${base}${n}@${domain}`;
    n++;
  }
  return email;
}

function generateStudentPassword() {
  return 'kshl' + String(Math.floor(100000 + Math.random() * 900000));
}

function addCreatedStudent(className, familyName, givenName, email, password) {
  createdStudentsStore.push({
    createdAt: new Date().toISOString(),
    className: className || '',
    familyName: familyName || '',
    givenName: givenName || '',
    email: email || '',
    password: password || ''
  });
  try {
    writeFileSync(CREATED_STUDENTS_FILE, JSON.stringify(createdStudentsStore, null, 2), 'utf8');
  } catch (_) {}
}

app.get('/api/auth/url', (req, res) => {
  try {
    const url = getAuthUrl();
    res.json({ url });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/auth/token', async (req, res) => {
  try {
    const { code } = req.body;
    const { tokens } = await getTokensFromCode(code);
    const user = await getUserInfo(tokens);
    res.json({ tokens, user });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/courses', async (req, res) => {
  try {
    const { tokens } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени' });
    const courses = await listCourses(tokens);
    res.json({ courses });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/roster', async (req, res) => {
  try {
    const { tokens, courseId } = req.body || {};
    if (!tokens || !courseId) return res.status(400).json({ error: 'Потрібні токени та courseId' });
    const [teachers, students] = await Promise.all([
      getCourseTeachers(tokens, courseId),
      getCourseStudents(tokens, courseId)
    ]);
    res.json({ teachers, students });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/add-student', async (req, res) => {
  try {
    const { tokens, courseId, email } = req.body || {};
    if (!tokens || !courseId || !email) return res.status(400).json({ error: 'Потрібні токени, courseId та email' });
    await addStudentToCourse(tokens, courseId, email);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/update', async (req, res) => {
  try {
    const { tokens, courseId, name, section } = req.body || {};
    if (!tokens || !courseId) return res.status(400).json({ error: 'Потрібні токени та courseId' });
    await updateCourse(tokens, courseId, name, section);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/remove-student', async (req, res) => {
  try {
    const { tokens, courseId, email } = req.body || {};
    if (!tokens || !courseId || !email) return res.status(400).json({ error: 'Потрібні токени, courseId та email' });
    await removeStudentFromCourse(tokens, courseId, email);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/add-teacher', async (req, res) => {
  try {
    const { tokens, courseId, email } = req.body || {};
    if (!tokens || !courseId || !email) return res.status(400).json({ error: 'Потрібні токени, courseId та email' });
    const classroom = getClassroomClient(tokens);
    await addTeacher(classroom, courseId, email);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/course/remove-teacher', async (req, res) => {
  try {
    const { tokens, courseId, email } = req.body || {};
    if (!tokens || !courseId || !email) return res.status(400).json({ error: 'Потрібні токени, courseId та email' });
    await removeTeacherFromCourse(tokens, courseId, email);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classroom/courses/archive', async (req, res) => {
  try {
    const { tokens, courseIds } = req.body || {};
    if (!tokens || !Array.isArray(courseIds) || !courseIds.length) return res.status(400).json({ error: 'Потрібні токени та courseIds' });
    const results = [];
    for (const courseId of courseIds) {
      try {
        await archiveCourse(tokens, courseId);
        results.push({ courseId, success: true });
      } catch (err) {
        results.push({ courseId, success: false, error: err.message });
      }
    }
    res.json({ results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/auth/user', async (req, res) => {
  try {
    const { tokens } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени' });
    const user = await getUserInfo(tokens);
    res.json({ user });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classes/admin', async (req, res) => {
  try {
    const { tokens, parentOuName } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени авторизації' });
    const classes = await getClassesFromAdmin(tokens, parentOuName || 'Учні');
    res.json({ classes });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classes/admin/stats', async (req, res) => {
  try {
    const { tokens, parentOuName } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени авторизації' });
    const classes = await getClassesFromAdmin(tokens, parentOuName || 'Учні');
    let studentsCount = 0;
    for (const c of classes) {
      if (c.orgUnitPath) {
        const users = await getUsersFromOrgUnit(tokens, c.orgUnitPath);
        studentsCount += users.length;
      }
    }
    res.json({ classesCount: classes.length, studentsCount });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classes/admin/students', async (req, res) => {
  try {
    const { tokens, orgUnitPath } = req.body || {};
    if (!tokens || !orgUnitPath) return res.status(400).json({ error: 'Потрібні токени та orgUnitPath' });
    const students = await getUsersFromOrgUnit(tokens, orgUnitPath);
    res.json({ students });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

function parseClassForPromotion(name) {
  const n = (name || '').trim();
  const m = n.match(/^(\d+)[-\s]*(.*)$/);
  if (m) return { grade: parseInt(m[1], 10), letter: (m[2] || '').trim() };
  return { grade: 0, letter: n };
}

function getPromotedClassName(name) {
  const { grade, letter } = parseClassForPromotion(name);
  if (grade >= 11) return null;
  const newGrade = grade + 1;
  return letter ? `${newGrade}-${letter}` : String(newGrade);
}

app.post('/api/classes/admin/rename', async (req, res) => {
  try {
    const { tokens, orgUnitPath, newName } = req.body || {};
    if (!tokens || !orgUnitPath || !newName) return res.status(400).json({ error: 'Потрібні токени, orgUnitPath та newName' });
    await updateOrgUnit(tokens, orgUnitPath, (newName || '').trim());
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classes/admin/delete', async (req, res) => {
  try {
    const { tokens, orgUnitPath } = req.body || {};
    if (!tokens || !orgUnitPath) return res.status(400).json({ error: 'Потрібні токени та orgUnitPath' });
    const users = await getUsersFromOrgUnit(tokens, orgUnitPath);
    for (const u of users || []) {
      const email = u.email || u.primaryEmail;
      if (email) await moveUserToOrgUnit(tokens, email, 'Випускники');
    }
    await deleteOrgUnit(tokens, orgUnitPath);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/classes/admin/promote', async (req, res) => {
  try {
    const { tokens } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени' });
    const classes = await getClassesFromAdmin(tokens, 'Учні');
    const grade11 = classes.filter(c => parseClassForPromotion(c.name).grade === 11);
    const others = classes.filter(c => {
      const g = parseClassForPromotion(c.name).grade;
      return g >= 1 && g <= 10;
    });
    const results = [];
    for (const c of grade11) {
      try {
        const users = await getUsersFromOrgUnit(tokens, c.orgUnitPath);
        for (const u of users || []) {
          const email = u.email || u.primaryEmail;
          if (email) {
            await moveUserToOrgUnit(tokens, email, 'Випускники');
          }
        }
        await deleteOrgUnit(tokens, c.orgUnitPath);
        results.push({ action: 'graduate', name: c.name, success: true });
      } catch (err) {
        results.push({ action: 'graduate', name: c.name, success: false, error: err.message });
      }
    }
    const sorted = others.sort((a, b) => {
      const ga = parseClassForPromotion(a.name).grade;
      const gb = parseClassForPromotion(b.name).grade;
      return gb - ga;
    });
    for (const c of sorted) {
      const newName = getPromotedClassName(c.name);
      if (!newName) continue;
      try {
        await updateOrgUnit(tokens, c.orgUnitPath, newName);
        results.push({ action: 'rename', oldName: c.name, newName, success: true });
      } catch (err) {
        results.push({ action: 'rename', oldName: c.name, newName, success: false, error: err.message });
      }
    }
    res.json({ results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/teachers/admin', async (req, res) => {
  try {
    const { tokens, ouName } = req.body || {};
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени авторизації' });
    const teachers = await getTeachersFromAdmin(tokens, ouName || 'Вчителі');
    res.json({ teachers });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/teachers/move-to-graduated', async (req, res) => {
  try {
    const { tokens, email } = req.body || {};
    if (!tokens || !email) return res.status(400).json({ error: 'Потрібні токени та email' });
    await moveUserToOrgUnit(tokens, email, 'Вибувші');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/teachers/create', async (req, res) => {
  try {
    const { tokens, givenName, familyName, primaryEmail, password } = req.body || {};
    if (!tokens || !givenName || !familyName || !primaryEmail || !password) {
      return res.status(400).json({ error: 'Потрібні токени, ім\'я, прізвище, email та пароль' });
    }
    await createUserInOrgUnit(tokens, { givenName, familyName, primaryEmail, password }, 'Вчителі');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users/search', async (req, res) => {
  try {
    const { tokens, searchTerm } = req.body || {};
    if (!tokens || !searchTerm) return res.status(400).json({ error: 'Потрібні токени та пошуковий запит' });
    const users = await searchUsersInDomain(tokens, searchTerm);
    res.json({ users });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users/create', async (req, res) => {
  try {
    const { tokens, givenName, familyName, primaryEmail, password, orgUnitPath } = req.body || {};
    if (!tokens || !givenName || !familyName || !primaryEmail || !password || !orgUnitPath) {
      return res.status(400).json({ error: 'Потрібні токени, ім\'я, прізвище, email, пароль та orgUnitPath' });
    }
    await createUserInOrgPath(tokens, { givenName, familyName, primaryEmail, password }, orgUnitPath);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/created-students/dates', (req, res) => {
  const dates = [...new Set(createdStudentsStore.map(s => s.createdAt.slice(0, 10)))].sort().reverse();
  res.json({ dates });
});

app.get('/api/created-students', (req, res) => {
  const { date } = req.query || {};
  let items = createdStudentsStore;
  if (date) items = items.filter(s => s.createdAt.slice(0, 10) === date);
  const byClass = {};
  for (const s of items) {
    const k = s.className || '(без класу)';
    if (!byClass[k]) byClass[k] = [];
    byClass[k].push({ familyName: s.familyName, givenName: s.givenName, email: s.email, password: s.password });
  }
  res.json({ byClass, items });
});

app.post('/api/created-students/create-and-add', async (req, res) => {
  try {
    const { tokens, studentName, className, courseId } = req.body || {};
    if (!tokens || !studentName || !courseId) return res.status(400).json({ error: 'Потрібні токени, studentName та courseId' });
    let adminClasses = [];
    try {
      adminClasses = await getClassesFromAdmin(tokens, 'Учні');
    } catch (_) {}
    const adminClass = findAdminClassForJournal(adminClasses, className || '');
    if (!adminClass?.orgUnitPath) return res.status(400).json({ error: 'Не знайдено клас у Google Admin для: ' + (className || '') });
    const { familyName, givenName } = parseStudentName(studentName);
    if (!familyName && !givenName) return res.status(400).json({ error: 'Невалідне ім\'я учня' });
    let classStudents = [];
    try {
      classStudents = await getUsersFromOrgUnit(tokens, adminClass.orgUnitPath);
    } catch (_) {}
    const existingEmails = new Set((classStudents || []).map(u => (u.email || u.primaryEmail || '').toLowerCase()).filter(Boolean));
    const pwd = generateStudentPassword();
    const email = generateStudentEmail(familyName, givenName, existingEmails);
    await createUserInOrgPath(tokens, { givenName, familyName, primaryEmail: email, password: pwd }, adminClass.orgUnitPath);
    const normClassName = normalizeClassForMatch(className) || className;
    addCreatedStudent(normClassName, familyName, givenName, email, pwd);
    await addStudentToCourse(tokens, courseId, email);
    res.json({ success: true, email, password: pwd, familyName, givenName });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users/move', async (req, res) => {
  try {
    const { tokens, email, orgUnitPath } = req.body || {};
    if (!tokens || !email || !orgUnitPath) return res.status(400).json({ error: 'Потрібні токени, email та orgUnitPath' });
    await moveUserToOrgPath(tokens, email, orgUnitPath);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users/move-to-graduates', async (req, res) => {
  try {
    const { tokens, email } = req.body || {};
    if (!tokens || !email) return res.status(400).json({ error: 'Потрібні токени та email' });
    await moveUserToOrgUnit(tokens, email, 'Випускники');
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/users/update', async (req, res) => {
  try {
    const { tokens, email, givenName, familyName, password } = req.body || {};
    if (!tokens || !email) return res.status(400).json({ error: 'Потрібні токени та email' });
    await updateUser(tokens, email, { givenName, familyName, password });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/teachers/map', (req, res) => {
  const { mapping } = req.body;
  if (Array.isArray(mapping)) {
    mapping.forEach(({ name, email }) => {
      if (name && email) teacherEmailMap.set(normalizeName(name), email);
    });
  } else if (typeof mapping === 'object') {
    Object.entries(mapping).forEach(([name, email]) => {
      if (name && email) teacherEmailMap.set(normalizeName(name), email);
    });
  }
  res.json({ ok: true });
});

function normalizeName(name) {
  return (name || '').replace(/[\u0027\u2019\u02BC\u0060\u00B4\u2032]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalizeNameForMatch(s) {
  return (s || '').replace(/[\u0027\u2019\u02BC\u0060\u00B4\u2032]/g, "'").replace(/\s+/g, ' ').trim().toLowerCase();
}

app.post('/api/parse', async (req, res) => {
  try {
    const { login, password } = req.body || {};
    const html = await fetchWithPuppeteer('https://nz.ua/journal/list', { login, password });
    const items = parseJournalList(html);
    if (items.length === 0 && (login || password)) {
      const hasJournalTable = html.includes('journal-choose') || html.includes('journal=');
      if (!hasJournalTable) {
        try {
          writeFileSync('debug-nz-page.html', html, 'utf8');
        } catch (_) {}
        return res.status(400).json({
          error: 'Не знайдено журнали. HTML збережено в debug-nz-page.html для перегляду. Можливо: 1) невірний логін/пароль 2) Cloudflare блокує 3) спробуйте headless: false в fetcher.js',
          items: []
        });
      }
    }
    res.json({ items });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/parse/journal', async (req, res) => {
  try {
    const { journalId, subgroupId, login, password } = req.body || {};
    if (!journalId || !login || !password) return res.status(400).json({ error: 'Потрібні journalId, login, password' });
    const url = `https://nz.ua/journal/index?journal=${journalId}${subgroupId ? `&subgroup=${subgroupId}` : ''}`;
    const html = await fetchWithPuppeteer(url, { login, password });
    const { teacher, students } = parseJournalDetails(html);
    res.json({ teacher, students });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/parse/teachers', async (req, res) => {
  try {
    const { journals, login, password } = req.body || {};
    const nzCreds = { login, password };
    const teachers = [];
    for (const j of journals || []) {
      const url = `https://nz.ua/journal/index?journal=${j.journalId}${j.subgroupId ? `&subgroup=${j.subgroupId}` : ''}`;
      const html = await fetchWithPuppeteer(url, nzCreds);
      const teacher = parseJournalPage(html);
      if (teacher) {
        teachers.push({
          ...teacher,
          journalId: j.journalId,
          subgroupId: j.subgroupId,
          classLabel: j.classLabel
        });
      }
    }
    res.json({ teachers });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

function buildTeacherMap(adminTeachers) {
  const map = new Map();
  for (const t of adminTeachers || []) {
    if (!t.email || !t.name) continue;
    const parts = t.name.trim().split(/\s+/).filter(Boolean);
    const keys = [
      t.name,
      parts.length >= 2 ? `${parts[parts.length - 1]} ${parts[0]}` : null,
      parts.length >= 2 ? `${parts[0]} ${parts[parts.length - 1]}` : null,
      parts.length >= 2 ? `${parts[parts.length - 1]} ${parts.slice(0, -1).join(' ')}` : null
    ].filter(Boolean);
    keys.forEach(k => map.set(normalizeName(k), t.email));
  }
  return map;
}

function findTeacherEmail(emailMap, nzName) {
  let key = normalizeName(nzName);
  if (emailMap.has(key)) return emailMap.get(key);
  const parts = nzName.trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) {
    key = normalizeName(`${parts[0]} ${parts[1]}`);
    if (emailMap.has(key)) return emailMap.get(key);
    key = normalizeName(`${parts[1]} ${parts[0]}`);
    if (emailMap.has(key)) return emailMap.get(key);
  }
  return null;
}

function normalizeClassForMatch(s) {
  if (!s || typeof s !== 'string') return '';
  let t = s.split('(')[0].trim().replace(/\s*клас\s*$/gi, '').trim();
  t = t.replace(/\s+/g, ' ').replace(/[\u2010-\u2015\u2212\-–—]/g, '-').trim();
  return t.replace(/-/g, '').replace(/\s+/g, '');
}

function findAdminClassForJournal(adminClasses, journalClassLabel) {
  const baseNorm = normalizeClassForMatch(journalClassLabel);
  if (!baseNorm) return null;
  const exact = (adminClasses || []).find(c => normalizeClassForMatch(c.name) === baseNorm);
  if (exact) return exact;
  const prefix = (adminClasses || []).find(c => {
    const adminNorm = normalizeClassForMatch(c.name);
    return adminNorm && (adminNorm === baseNorm || adminNorm.startsWith(baseNorm) || baseNorm.startsWith(adminNorm));
  });
  return prefix || null;
}

function findStudentEmail(adminStudents, studentName) {
  const parts = (studentName || '').trim().split(/\s+/).filter(Boolean);
  const targetKey = normalizeNameForMatch(studentName);
  const targetSwapped = parts.length >= 2 ? normalizeNameForMatch(parts.slice(1).concat(parts[0]).join(' ')) : targetKey;
  const targetFirstTwo = parts.length >= 2 ? normalizeNameForMatch(parts.slice(0, 2).join(' ')) : targetKey;
  const targetFirstTwoSwapped = parts.length >= 2 ? normalizeNameForMatch(parts[1] + ' ' + parts[0]) : targetKey;
  for (const u of adminStudents || []) {
    const fn = (u.familyName || '').trim();
    const gn = (u.givenName || '').trim();
    const key1 = normalizeNameForMatch(`${fn} ${gn}`);
    const key2 = normalizeNameForMatch(`${gn} ${fn}`);
    if (key1 === targetKey || key2 === targetKey || key1 === targetSwapped || key2 === targetSwapped) return u.email;
    if (key1 === targetFirstTwo || key2 === targetFirstTwo || key1 === targetFirstTwoSwapped || key2 === targetFirstTwoSwapped) return u.email;
    if (targetKey.startsWith(key1 + ' ') || targetKey.startsWith(key2 + ' ')) return u.email;
    if (targetSwapped.startsWith(key1 + ' ') || targetSwapped.startsWith(key2 + ' ')) return u.email;
  }
  return null;
}

app.post('/api/sync/init', async (req, res) => {
  try {
    const { tokens, teacherMapping, login, password, ouName } = req.body;
    if (!tokens || !login || !password) return res.status(400).json({ error: 'Потрібні токени, логін та пароль' });
    let currentUserEmail = null;
    try {
      const user = await getUserInfo(tokens);
      currentUserEmail = user?.email;
    } catch (_) {}
    let emailMap = teacherEmailMap;
    try {
      const adminTeachers = await getTeachersFromAdmin(tokens, ouName || 'Вчителі');
      emailMap = buildTeacherMap(adminTeachers);
    } catch (_) {}
    if (teacherMapping && Array.isArray(teacherMapping)) {
      teacherMapping.forEach(({ name, email }) => {
        if (name && email) emailMap.set(normalizeName(name), email);
      });
    }
    let adminClasses = [];
    try {
      adminClasses = await getClassesFromAdmin(tokens, 'Учні');
    } catch (_) {}
    const nzCreds = { login, password };
    const html = await fetchWithPuppeteer('https://nz.ua/journal/list', nzCreds);
    const items = parseJournalList(html);
    const journals = [];
    for (const { subject, classes } of items) {
      for (const cls of classes) {
        journals.push({
          subject,
          classLabel: cls.classLabel,
          journalId: cls.journalId,
          subgroupId: cls.subgroupId
        });
      }
    }
    const emailMapArr = [...emailMap.entries()];
    res.json({ journals, emailMap: emailMapArr, adminClasses, currentUserEmail });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

function findExistingCourse(existingCourses, section, name) {
  const s = (section || '').trim();
  const n = (name || '').trim();
  return (existingCourses || []).find(c => (c.section || '').trim() === s && (c.name || '').trim() === n);
}

app.post('/api/sync/create-one', async (req, res) => {
  try {
    const { tokens, login, password, journal, emailMapArr, adminClasses, currentUserEmail } = req.body;
    if (!tokens || !login || !password || !journal) return res.status(400).json({ error: 'Потрібні токени, логін, пароль та journal' });
    const emailMap = new Map(emailMapArr || []);
    const classroom = getClassroomClient(tokens);
    const nzCreds = { login, password };
    const courseName = `${journal.subject} - ${journal.classLabel}`;
    try {
      const existingCourses = await listCourses(tokens);
      let course = findExistingCourse(existingCourses, journal.subject, journal.classLabel);
      if (!course) course = await createCourse(classroom, journal.classLabel, 'me', journal.subject);
      const cacheKey = `${journal.journalId}-${journal.subgroupId || ''}`;
      const journalUrl = `https://nz.ua/journal/index?journal=${journal.journalId}${journal.subgroupId ? `&subgroup=${journal.subgroupId}` : ''}`;
      const journalHtml = await fetchWithPuppeteer(journalUrl, nzCreds);
      const { teacher, students } = parseJournalDetails(journalHtml);
      const teacherEmail = teacher ? findTeacherEmail(emailMap, teacher.name) : null;
      const teachersAdded = new Set();
      if (currentUserEmail && teacherEmail && teacherEmail.toLowerCase() === currentUserEmail.toLowerCase()) {
        teachersAdded.add(currentUserEmail);
      }
      if (teacherEmail && !teachersAdded.has(teacherEmail)) {
        await addTeacher(classroom, course.id, teacherEmail);
        teachersAdded.add(teacherEmail);
      }
      let studentsAdded = 0;
      const adminClass = findAdminClassForJournal(adminClasses, journal.classLabel);
      let classStudents = [];
      if (adminClass?.orgUnitPath) {
        try {
          classStudents = await getUsersFromOrgUnit(tokens, adminClass.orgUnitPath);
        } catch (_) {}
      }
      const existingEmails = new Set((classStudents || []).map(u => (u.email || u.primaryEmail || '').toLowerCase()).filter(Boolean));
      const className = normalizeClassForMatch(journal.classLabel) || journal.classLabel;
      for (const st of (students || [])) {
        let email = findStudentEmail(classStudents, st.name);
        if (!email && adminClass?.orgUnitPath) {
          const { familyName, givenName } = parseStudentName(st.name);
          if (familyName || givenName) {
            try {
              const pwd = generateStudentPassword();
              email = generateStudentEmail(familyName, givenName, existingEmails);
              await createUserInOrgPath(tokens, { givenName, familyName, primaryEmail: email, password: pwd }, adminClass.orgUnitPath);
              existingEmails.add(email.toLowerCase());
              classStudents.push({ familyName, givenName, email, primaryEmail: email });
              addCreatedStudent(className, familyName, givenName, email, pwd);
            } catch (createErr) {}
          }
        }
        if (email) {
          try {
            await addStudentToCourse(tokens, course.id, email);
            studentsAdded++;
          } catch (_) {}
        }
      }
      res.json({ courseName, courseId: course.id, teacherAdded: teachersAdded.size > 0, studentsAdded });
    } catch (err) {
      res.json({ courseName, error: err.message });
    }
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/sync', async (req, res) => {
  try {
    const { tokens, teacherMapping, login, password, ouName } = req.body;
    if (!tokens) return res.status(400).json({ error: 'Потрібні токени авторизації' });

    let currentUserEmail = null;
    try {
      const user = await getUserInfo(tokens);
      currentUserEmail = user?.email;
    } catch (_) {}

    let emailMap = teacherEmailMap;
    try {
      const adminTeachers = await getTeachersFromAdmin(tokens, ouName || 'Вчителі');
      emailMap = buildTeacherMap(adminTeachers);
    } catch (_) {}

    if (teacherMapping && Array.isArray(teacherMapping)) {
      teacherMapping.forEach(({ name, email }) => {
        if (name && email) emailMap.set(normalizeName(name), email);
      });
    }

    let adminClasses = [];
    try {
      adminClasses = await getClassesFromAdmin(tokens, 'Учні');
    } catch (_) {}

    const classStudentsCache = new Map();
    const nzCreds = { login, password };
    const html = await fetchWithPuppeteer('https://nz.ua/journal/list', nzCreds);
    const items = parseJournalList(html);
    const classroom = getClassroomClient(tokens);
    const results = [];
    const allJournals = [];

    for (const { subject, classes } of items) {
      for (const cls of classes) {
        allJournals.push({
          subject,
          classLabel: cls.classLabel,
          journalId: cls.journalId,
          subgroupId: cls.subgroupId
        });
      }
    }

    const journalCache = new Map();
    let existingCourses = await listCourses(tokens);
    for (let i = 0; i < allJournals.length; i++) {
      const j = allJournals[i];
      const courseName = `${j.subject} - ${j.classLabel}`;
      try {
        let course = findExistingCourse(existingCourses, j.subject, j.classLabel);
        if (!course) {
          course = await createCourse(classroom, j.classLabel, 'me', j.subject);
          existingCourses = existingCourses.concat([{ id: course.id, name: course.name || j.classLabel, section: course.section || j.subject }]);
        }
        const cacheKey = `${j.journalId}-${j.subgroupId || ''}`;
        let teacherEmail = null;
        let students = [];
        if (!journalCache.has(cacheKey)) {
          const journalUrl = `https://nz.ua/journal/index?journal=${j.journalId}${j.subgroupId ? `&subgroup=${j.subgroupId}` : ''}`;
          const journalHtml = await fetchWithPuppeteer(journalUrl, nzCreds);
          const { teacher, students: st } = parseJournalDetails(journalHtml);
          journalCache.set(cacheKey, { teacher, students: st || [] });
          teacherEmail = teacher ? findTeacherEmail(emailMap, teacher.name) : null;
          students = st || [];
        } else {
          const cached = journalCache.get(cacheKey);
          teacherEmail = cached.teacher ? findTeacherEmail(emailMap, cached.teacher.name) : null;
          students = cached.students || [];
        }
        const teachersAdded = new Set();
        if (teacherEmail && currentUserEmail && teacherEmail.toLowerCase() === currentUserEmail.toLowerCase()) {
          teachersAdded.add(currentUserEmail);
        }
        if (teacherEmail && !teachersAdded.has(teacherEmail)) {
          await addTeacher(classroom, course.id, teacherEmail);
          teachersAdded.add(teacherEmail);
        }
        let studentsAdded = 0;
        const adminClass = findAdminClassForJournal(adminClasses, j.classLabel);
        let classStudents = [];
        if (adminClass?.orgUnitPath) {
          if (!classStudentsCache.has(adminClass.orgUnitPath)) {
            try {
              classStudents = await getUsersFromOrgUnit(tokens, adminClass.orgUnitPath);
              classStudentsCache.set(adminClass.orgUnitPath, classStudents);
            } catch (_) {}
          } else {
            classStudents = classStudentsCache.get(adminClass.orgUnitPath) || [];
          }
        }
        const existingEmails = new Set((classStudents || []).map(u => (u.email || u.primaryEmail || '').toLowerCase()).filter(Boolean));
        const className = normalizeClassForMatch(j.classLabel) || j.classLabel;
        for (const st of students) {
          let email = findStudentEmail(classStudents, st.name);
          if (!email && adminClass?.orgUnitPath) {
            const { familyName, givenName } = parseStudentName(st.name);
            if (familyName || givenName) {
              try {
                const pwd = generateStudentPassword();
                email = generateStudentEmail(familyName, givenName, existingEmails);
                await createUserInOrgPath(tokens, { givenName, familyName, primaryEmail: email, password: pwd }, adminClass.orgUnitPath);
                existingEmails.add(email.toLowerCase());
                classStudents.push({ familyName, givenName, email, primaryEmail: email });
                classStudentsCache.set(adminClass.orgUnitPath, classStudents);
                addCreatedStudent(className, familyName, givenName, email, pwd);
              } catch (createErr) {}
            }
          }
          if (email) {
            try {
              await addStudentToCourse(tokens, course.id, email);
              studentsAdded++;
            } catch (_) {}
          }
        }
        results.push({ courseName, courseId: course.id, teacherAdded: teachersAdded.size > 0, studentsAdded });
      } catch (err) {
        results.push({ courseName, error: err.message });
      }
    }
    res.json({ results });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const distPath = join(__dirname, '..', 'dist');
app.use(express.static(distPath));
app.get('*', (req, res) => {
  res.sendFile(join(distPath, 'index.html'));
});

const PORT = process.env.PORT || 3003;
app.listen(PORT, () => console.log(`Server: http://localhost:${PORT}`));
