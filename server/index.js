import 'dotenv/config';
import express from 'express';
import { writeFileSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';
import cors from 'cors';
import { fetchWithPuppeteer } from './fetcher.js';
import { parseJournalList, parseJournalPage, parseJournalDetails } from './parser.js';
import { getClassroomClient, createCourse, addTeacher, addStudentToCourse, removeStudentFromCourse, removeTeacherFromCourse, archiveCourse, getAuthUrl, getTokensFromCode, getTeachersFromAdmin, getClassesFromAdmin, getUsersFromOrgUnit, getUserInfo, moveUserToOrgUnit, createUserInOrgUnit, searchUsersInDomain, createUserInOrgPath, moveUserToOrgPath, updateUser, listCourses, getCourseTeachers, getCourseStudents } from './classroom.js';

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
app.use(cors());
app.use(express.json());

const teacherEmailMap = new Map();

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
  return name.replace(/\s+/g, ' ').trim().toLowerCase();
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
  return s.split('(')[0].trim().replace(/\s+/g, ' ').replace(/-/g, '');
}

function findStudentEmail(adminStudents, studentName) {
  const key = normalizeName(studentName);
  for (const u of adminStudents || []) {
    const n = normalizeName([u.givenName, u.familyName].filter(Boolean).join(' '));
    if (n === key) return u.email;
    const n2 = normalizeName([u.familyName, u.givenName].filter(Boolean).join(' '));
    if (n2 === key) return u.email;
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

app.post('/api/sync/create-one', async (req, res) => {
  try {
    const { tokens, login, password, journal, emailMapArr, adminClasses, currentUserEmail } = req.body;
    if (!tokens || !login || !password || !journal) return res.status(400).json({ error: 'Потрібні токени, логін, пароль та journal' });
    const emailMap = new Map(emailMapArr || []);
    const classroom = getClassroomClient(tokens);
    const nzCreds = { login, password };
    const courseName = `${journal.subject} - ${journal.classLabel}`;
    try {
      const course = await createCourse(classroom, journal.classLabel, 'me', journal.subject);
      const cacheKey = `${journal.journalId}-${journal.subgroupId || ''}`;
      const journalUrl = `https://nz.ua/journal/index?journal=${journal.journalId}${journal.subgroupId ? `&subgroup=${journal.subgroupId}` : ''}`;
      const journalHtml = await fetchWithPuppeteer(journalUrl, nzCreds);
      const { teacher, students } = parseJournalDetails(journalHtml);
      const teacherEmail = teacher ? findTeacherEmail(emailMap, teacher.name) : null;
      const teachersAdded = new Set();
      if (currentUserEmail) {
        await addTeacher(classroom, course.id, currentUserEmail);
        teachersAdded.add(currentUserEmail);
      }
      if (teacherEmail && !teachersAdded.has(teacherEmail)) {
        await addTeacher(classroom, course.id, teacherEmail);
        teachersAdded.add(teacherEmail);
      }
      let studentsAdded = 0;
      const normClass = normalizeClassForMatch(journal.classLabel);
      const adminClass = (adminClasses || []).find(c => normalizeClassForMatch(c.name) === normClass);
      let classStudents = [];
      if (adminClass?.orgUnitPath) {
        try {
          classStudents = await getUsersFromOrgUnit(tokens, adminClass.orgUnitPath);
        } catch (_) {}
      }
      for (const st of (students || [])) {
        const email = findStudentEmail(classStudents, st.name);
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
    for (let i = 0; i < allJournals.length; i++) {
      const j = allJournals[i];
      const courseName = `${j.subject} - ${j.classLabel}`;
      try {
        const course = await createCourse(classroom, j.classLabel, 'me', j.subject);
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
        if (currentUserEmail) {
          await addTeacher(classroom, course.id, currentUserEmail);
          teachersAdded.add(currentUserEmail);
        }
        if (teacherEmail && !teachersAdded.has(teacherEmail)) {
          await addTeacher(classroom, course.id, teacherEmail);
          teachersAdded.add(teacherEmail);
        }
        let studentsAdded = 0;
        const normClass = normalizeClassForMatch(j.classLabel);
        const adminClass = adminClasses.find(c => normalizeClassForMatch(c.name) === normClass);
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
        for (const st of students) {
          const email = findStudentEmail(classStudents, st.name);
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
