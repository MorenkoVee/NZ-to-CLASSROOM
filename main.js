const API = '/api';
const UA_TO_LATIN = { 'а':'a','б':'b','в':'v','г':'h','ґ':'g','д':'d','е':'e','є':'ie','ж':'zh','з':'z','и':'y','і':'i','ї':'i','й':'i','к':'k','л':'l','м':'m','н':'n','о':'o','п':'p','р':'r','с':'s','т':'t','у':'u','ф':'f','х':'kh','ц':'ts','ч':'ch','ш':'sh','щ':'shch','ь':'','ю':'iu','я':'ia','ё':'io' };
const CACHE_KEYS = { parsedItems: 'classroom_parsed_items', teachers: 'classroom_teachers', adminStats: 'classroom_admin_stats', adminClasses: 'classroom_admin_classes', classroomCourses: 'classroom_classroom_courses', courseRosters: 'classroom_course_rosters', journalDetails: 'classroom_journal_details', classStudents: 'classroom_class_students' };

let tokens = null;
let parsedItems = [];
let teachers = [];
let adminClassesCount = 0;
let adminStudentsCount = 0;
let adminClasses = [];
let classroomCourses = [];
const journalLoading = new Set();
const classStudentsLoading = new Set();
const courseRosterLoading = new Set();
const courseRosterCache = {};

function transliterateUAtoEN(s) {
  return (s || '').toLowerCase().split('').map(c => {
    const lower = c.toLowerCase();
    if (UA_TO_LATIN[lower]) return UA_TO_LATIN[lower];
    if (/[a-z0-9]/.test(c)) return c;
    return '';
  }).join('');
}

function generateTeacherEmail(familyName, givenName) {
  const fn = transliterateUAtoEN((familyName || '').replace(/\s+/g, ' ').trim());
  const gn = transliterateUAtoEN((givenName || '').replace(/\s+/g, ' ').trim());
  const part = [fn, gn].filter(Boolean).join('_').replace(/[^a-z0-9_]/g, '') || 'user';
  return `${part}@kshg.site`;
}

function generateTeacherPassword() {
  return 'kshl' + String(Math.floor(100000 + Math.random() * 900000));
}

function loadTokens() {
  const s = localStorage.getItem('classroom_tokens');
  if (s) {
    try {
      tokens = JSON.parse(s);
      return true;
    } catch (_) {}
  }
  return false;
}

function saveTokens(t, user) {
  tokens = t;
  localStorage.setItem('classroom_tokens', JSON.stringify(t));
  if (user) localStorage.setItem('classroom_user', JSON.stringify(user));
}

function loadFromCache() {
  try {
    const pi = localStorage.getItem(CACHE_KEYS.parsedItems);
    const t = localStorage.getItem(CACHE_KEYS.teachers);
    const st = localStorage.getItem(CACHE_KEYS.adminStats);
    const ac = localStorage.getItem(CACHE_KEYS.adminClasses);
    const cc = localStorage.getItem(CACHE_KEYS.classroomCourses);
    return {
      parsedItems: pi ? JSON.parse(pi) : null,
      teachers: t ? JSON.parse(t) : null,
      adminStats: st ? JSON.parse(st) : null,
      adminClasses: ac ? JSON.parse(ac) : null,
      classroomCourses: cc ? JSON.parse(cc) : null
    };
  } catch (_) { return {}; }
}

function saveToCache() {
  try {
    if (parsedItems.length) localStorage.setItem(CACHE_KEYS.parsedItems, JSON.stringify(parsedItems));
    if (teachers.length) localStorage.setItem(CACHE_KEYS.teachers, JSON.stringify(teachers));
    localStorage.setItem(CACHE_KEYS.adminStats, JSON.stringify({ adminClassesCount, adminStudentsCount }));
    if (adminClasses.length) localStorage.setItem(CACHE_KEYS.adminClasses, JSON.stringify(adminClasses));
    if (classroomCourses.length) localStorage.setItem(CACHE_KEYS.classroomCourses, JSON.stringify(classroomCourses));
  } catch (_) {}
}

function getJournalDetailsCache(journalId, subgroupId) {
  try {
    const key = `${journalId}-${subgroupId || ''}`;
    const all = localStorage.getItem(CACHE_KEYS.journalDetails);
    if (!all) return null;
    const obj = JSON.parse(all);
    return obj[key] || null;
  } catch (_) { return null; }
}

function setJournalDetailsCache(journalId, subgroupId, data) {
  try {
    const key = `${journalId}-${subgroupId || ''}`;
    const all = localStorage.getItem(CACHE_KEYS.journalDetails) || '{}';
    const obj = JSON.parse(all);
    obj[key] = data;
    localStorage.setItem(CACHE_KEYS.journalDetails, JSON.stringify(obj));
  } catch (_) {}
}

function getClassStudentsCache(orgUnitPath) {
  try {
    const all = localStorage.getItem(CACHE_KEYS.classStudents);
    if (!all) return null;
    const obj = JSON.parse(all);
    return obj[orgUnitPath] || null;
  } catch (_) { return null; }
}

function setClassStudentsCache(orgUnitPath, students) {
  try {
    const all = localStorage.getItem(CACHE_KEYS.classStudents) || '{}';
    const obj = JSON.parse(all);
    obj[orgUnitPath] = students;
    localStorage.setItem(CACHE_KEYS.classStudents, JSON.stringify(obj));
  } catch (_) {}
}

function getCourseRosterCache(courseId) {
  try {
    const all = localStorage.getItem(CACHE_KEYS.courseRosters);
    if (!all) return null;
    const obj = JSON.parse(all);
    return obj[courseId] || null;
  } catch (_) { return null; }
}

function setCourseRosterCache(courseId, data) {
  try {
    const all = localStorage.getItem(CACHE_KEYS.courseRosters) || '{}';
    const obj = JSON.parse(all);
    obj[courseId] = data;
    localStorage.setItem(CACHE_KEYS.courseRosters, JSON.stringify(obj));
  } catch (_) {}
}

function clearGoogleCache() {
  [CACHE_KEYS.teachers, CACHE_KEYS.adminStats, CACHE_KEYS.adminClasses, CACHE_KEYS.classroomCourses, CACHE_KEYS.courseRosters, CACHE_KEYS.classStudents].forEach(k => localStorage.removeItem(k));
  journalLoading.clear();
  classStudentsLoading.clear();
  courseRosterLoading.clear();
  Object.keys(courseRosterCache).forEach(k => delete courseRosterCache[k]);
  classroomCourses = [];
  teachers = [];
  adminClassesCount = 0;
  adminStudentsCount = 0;
  adminClasses = [];
  updateStats();
  const listEl = document.getElementById('adminTeachersList');
  if (listEl) { listEl.style.display = 'none'; listEl.innerHTML = ''; }
  const journalTeachersEl = document.getElementById('journalTeachersList');
  if (journalTeachersEl) journalTeachersEl.innerHTML = '';
  const cardJournalTeachers = document.getElementById('cardTeachersJournals');
  if (cardJournalTeachers) cardJournalTeachers.style.display = 'none';
  const classesEl = document.getElementById('adminClassesList');
  if (classesEl) { classesEl.style.display = 'none'; classesEl.innerHTML = ''; }
  const classroomEl = document.getElementById('classroomCoursesList');
  if (classroomEl) { classroomEl.style.display = 'none'; classroomEl.innerHTML = ''; }
  const classroomMass = document.getElementById('classroomCoursesMassActions');
  if (classroomMass) { classroomMass.style.display = 'none'; }
  const classroomStatus = document.getElementById('classroomCoursesStatus');
  if (classroomStatus) { classroomStatus.style.display = 'none'; classroomStatus.textContent = ''; }
  const preloadClassroom = document.getElementById('preloadClassroomStatus');
  if (preloadClassroom) { preloadClassroom.style.display = 'none'; preloadClassroom.textContent = ''; }
  const preloadClassroomDash = document.getElementById('preloadClassroomStatusDashboard');
  if (preloadClassroomDash) { preloadClassroomDash.style.display = 'none'; preloadClassroomDash.textContent = ''; }
  const preloadClasses = document.getElementById('preloadClassesStatus');
  if (preloadClasses) { preloadClasses.style.display = 'none'; preloadClasses.textContent = ''; }
  const preloadClassesDash = document.getElementById('preloadClassesStatusDashboard');
  if (preloadClassesDash) { preloadClassesDash.style.display = 'none'; preloadClassesDash.textContent = ''; }
  if (parsedItems.length) {
    const table = document.getElementById('parseTable');
    if (table) { table.style.display = 'block'; renderJournalsTable(); }
    renderJournalTeachersTable();
    const cardJournalTeachers = document.getElementById('cardTeachersJournals');
    if (cardJournalTeachers) cardJournalTeachers.style.display = 'block';
  }
}

function clearCache() {
  Object.values(CACHE_KEYS).forEach(k => localStorage.removeItem(k));
  journalLoading.clear();
  classStudentsLoading.clear();
  courseRosterLoading.clear();
  Object.keys(courseRosterCache).forEach(k => delete courseRosterCache[k]);
  parsedItems = [];
  classroomCourses = [];
  teachers = [];
  adminClassesCount = 0;
  adminStudentsCount = 0;
  adminClasses = [];
  updateStats();
  const table = document.getElementById('parseTable');
  if (table) { table.style.display = 'none'; table.innerHTML = ''; }
  const listEl = document.getElementById('adminTeachersList');
  if (listEl) { listEl.style.display = 'none'; listEl.innerHTML = ''; }
  const journalTeachersEl = document.getElementById('journalTeachersList');
  if (journalTeachersEl) journalTeachersEl.innerHTML = '';
  const cardJournalTeachers = document.getElementById('cardTeachersJournals');
  if (cardJournalTeachers) cardJournalTeachers.style.display = 'none';
  const classesEl = document.getElementById('adminClassesList');
  if (classesEl) { classesEl.style.display = 'none'; classesEl.innerHTML = ''; }
  const classroomEl = document.getElementById('classroomCoursesList');
  if (classroomEl) { classroomEl.style.display = 'none'; classroomEl.innerHTML = ''; }
  const classroomMass = document.getElementById('classroomCoursesMassActions');
  if (classroomMass) { classroomMass.style.display = 'none'; }
  const classroomStatus = document.getElementById('classroomCoursesStatus');
  if (classroomStatus) { classroomStatus.style.display = 'none'; classroomStatus.textContent = ''; }
  const preloadClassroom = document.getElementById('preloadClassroomStatus');
  if (preloadClassroom) { preloadClassroom.style.display = 'none'; preloadClassroom.textContent = ''; }
  const preloadClassroomDash = document.getElementById('preloadClassroomStatusDashboard');
  if (preloadClassroomDash) { preloadClassroomDash.style.display = 'none'; preloadClassroomDash.textContent = ''; }
  const status = document.getElementById('parseStatus');
  if (status) { status.style.display = 'none'; status.textContent = ''; }
  const preload = document.getElementById('preloadStatus');
  if (preload) { preload.style.display = 'none'; preload.textContent = ''; }
  const preloadDash = document.getElementById('preloadStatusDashboard');
  if (preloadDash) { preloadDash.style.display = 'none'; preloadDash.textContent = ''; }
  const preloadClasses = document.getElementById('preloadClassesStatus');
  if (preloadClasses) { preloadClasses.style.display = 'none'; preloadClasses.textContent = ''; }
  const preloadClassesDash = document.getElementById('preloadClassesStatusDashboard');
  if (preloadClassesDash) { preloadClassesDash.style.display = 'none'; preloadClassesDash.textContent = ''; }
}

const panelTitles = {
  dashboard: 'Головна',
  journals: 'Журнали',
  teachers: 'Вчителі',
  classes: 'Класи',
  classroom: 'Classroom',
  manage: 'Керування'
};

function showPanel(id) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const panel = document.getElementById('panel-' + id);
  const nav = document.querySelector('[data-panel="' + id + '"]');
  if (panel) panel.classList.add('active');
  if (nav) nav.classList.add('active');
  const title = document.getElementById('headerTitle');
  if (title) title.textContent = panelTitles[id] || id;
  if (id === 'dashboard') renderDashboardMismatches();
  if (id === 'manage') updateManagePanelState();
}

function updateManagePanelState() {
  const warnEl = document.getElementById('manageMismatchWarn');
  const btn = document.getElementById('btnSync');
  if (!warnEl || !btn) return;
  const { adminMismatches } = getMismatchClasses();
  if (adminMismatches.length > 0) {
    warnEl.style.display = 'block';
    warnEl.textContent = 'У Класах з Google Admin та Вчителі не має бути невідповідностей. Виправте невідповідності на Головній перед створенням курсів.';
    btn.disabled = true;
  } else {
    warnEl.style.display = 'none';
    btn.disabled = false;
  }
}

function updateAuthUI(user) {
  const status = document.getElementById('authStatus');
  const login = document.getElementById('authLogin');
  const userEl = document.getElementById('authUser');
  if (user) {
    userEl.textContent = user.email || user.name || 'Увійшов';
    status.style.display = 'flex';
    status.style.gap = '0.5rem';
    status.style.alignItems = 'center';
    login.style.display = 'none';
  } else {
    status.style.display = 'none';
    login.style.display = 'block';
  }
}

document.getElementById('btnAuth').onclick = async () => {
  try {
    const r = await fetch(`${API}/auth/url`);
    const { url } = await r.json();
    window.location.href = url;
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  }
};

if (window.location.pathname === '/callback' || window.location.search.includes('code=')) {
  const params = new URLSearchParams(window.location.search);
  const code = params.get('code');
  if (code) {
    (async () => {
      try {
        const r = await fetch(`${API}/auth/token`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ code })
        });
        const data = await r.json();
        if (data.tokens) {
          saveTokens(data.tokens, data.user);
          window.location.href = '/';
        } else throw new Error(data.error || 'Невідома помилка');
      } catch (e) {
        showAlertModal('Помилка авторизації: ' + e.message, 'Помилка', true);
      }
    })();
  }
}

document.getElementById('btnLogout').onclick = () => {
  tokens = null;
  localStorage.removeItem('classroom_tokens');
  localStorage.removeItem('classroom_user');
  clearGoogleCache();
  updateAuthUI(null);
};

(async function initAuth() {
  if (loadTokens()) {
    const cache = loadFromCache();
    if (cache.classroomCourses?.length) {
      classroomCourses = cache.classroomCourses;
      classroomCourses.forEach(c => {
        const id = c.id || c.courseId;
        const roster = getCourseRosterCache(id);
        if (roster) courseRosterCache[id] = roster;
      });
      renderClassroomTable();
      renderDashboardMismatches();
    }
    const cached = localStorage.getItem('classroom_user');
    if (cached) {
      try {
        updateAuthUI(JSON.parse(cached));
        return;
      } catch (_) {}
    }
    try {
      const r = await fetch(`${API}/auth/user`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens })
      });
      const data = await r.json();
      if (data.user) {
        saveTokens(tokens, data.user);
        updateAuthUI(data.user);
      }
    } catch (_) {}
  }
})();

function getNzCreds() {
  let login = document.getElementById('nzLogin')?.value?.trim();
  let password = document.getElementById('nzPassword')?.value?.trim();
  if (!login) login = localStorage.getItem('nz_login');
  if (!password) password = localStorage.getItem('nz_password');
  return { login: login || '', password: password || '' };
}

function saveNzCreds(login, password) {
  if (login) localStorage.setItem('nz_login', login);
  if (password) localStorage.setItem('nz_password', password);
  updateNzAuthUI();
}

function updateNzAuthUI() {
  const login = localStorage.getItem('nz_login');
  const password = localStorage.getItem('nz_password');
  const status = document.getElementById('nzAuthStatus');
  const form = document.getElementById('nzAuthLogin');
  const userEl = document.getElementById('nzAuthUser');
  if (login && password) {
    if (userEl) userEl.textContent = login;
    if (status) { status.style.display = 'flex'; status.style.gap = '0.5rem'; status.style.alignItems = 'center'; }
    if (form) form.style.display = 'none';
  } else {
    if (status) status.style.display = 'none';
    if (form) form.style.display = 'flex';
  }
}

function areAllJournalDetailsLoaded() {
  const journalRows = parsedItems.flatMap(({ classes }) => classes.map(c => ({ journalId: c.journalId, subgroupId: c.subgroupId })));
  if (!journalRows.length) return true;
  const cached = journalRows.filter(j => getJournalDetailsCache(j.journalId, j.subgroupId)).length;
  return cached >= journalRows.length;
}

function getJournalStats() {
  if (!areAllJournalDetailsLoaded()) return { teachersCount: null, studentsCount: null };
  let teachersCount = 0;
  let studentsCount = 0;
  try {
    const all = localStorage.getItem(CACHE_KEYS.journalDetails);
    if (all) {
      const obj = JSON.parse(all);
      const teacherNames = new Set();
      const studentIds = new Set();
      Object.values(obj).forEach(d => {
        if (d?.teacher?.name) teacherNames.add(d.teacher.name);
        (d?.students || []).forEach(s => {
          if (s?.studentId) studentIds.add(s.studentId);
          else if (s?.name) studentIds.add(s.name);
        });
      });
      teachersCount = teacherNames.size;
      studentsCount = studentIds.size;
    }
  } catch (_) {}
  return { teachersCount, studentsCount };
}

function normalizeClassLabel(label) {
  if (!label || typeof label !== 'string') return '';
  let s = label.split('(')[0].trim();
  s = s.replace(/\s*клас\s*$/gi, '').trim();
  s = s.replace(/-\s*$/, '').trim();
  return s;
}

function normalizeSubject(s) {
  if (!s || typeof s !== 'string') return '';
  return s.replace(/\s+/g, ' ').replace(/\s*,\s*/g, ', ').trim();
}

function classLabelMatches(section, journalClassLabel) {
  const s = (section || '').trim();
  const j = (journalClassLabel || '').trim();
  if (!s && !j) return true;
  if (!s || !j) return false;
  const normS = normalizeClassLabel(s);
  const normJ = normalizeClassLabel(j);
  if (normS !== normJ) return false;
  const extractSub = (x) => (x.match(/\(([^)]+)\)/) || [])[1] || '';
  const subS = extractSub(s).replace(/\s+/g, ' ').toLowerCase();
  const subJ = extractSub(j).replace(/\s+/g, ' ').toLowerCase();
  if (!subS && !subJ) return true;
  if (!subS || !subJ) return subS === subJ;
  const getSubNum = (x) => {
    if (/1|і\b|перш/i.test(x)) return 1;
    if (/2|іі|друг/i.test(x)) return 2;
    if (/3|ііі|трет/i.test(x)) return 3;
    const m = x.match(/\d/);
    return m ? parseInt(m[0], 10) : 0;
  };
  return getSubNum(subS) === getSubNum(subJ);
}

function looksLikeClassLabel(s) {
  return /^\d/.test((s || '').trim());
}

function parseCourseForJournal(courseName, courseSection) {
  const name = (courseName || '').trim();
  const section = (courseSection || '').trim();
  if (section && looksLikeClassLabel(name) && !looksLikeClassLabel(section)) {
    return { subject: section, section: name };
  }
  const sep = name.match(/\s+[-–—]\s+/);
  if (sep) {
    const idx = name.search(/\s+[-–—]\s+/);
    const part1 = name.slice(0, idx).trim();
    const part2 = name.slice(idx + sep[0].length).trim();
    if (looksLikeClassLabel(part2)) return { subject: part1, section: part2 };
    return { subject: part1, section: section || part2 };
  }
  if (looksLikeClassLabel(section) && !looksLikeClassLabel(name)) {
    return { subject: name, section };
  }
  return { subject: section || name, section: name || section };
}

function getStudentsFromJournalsForCourse(subject, section) {
  const parsed = parseCourseForJournal(subject, section);
  const normSubject = normalizeSubject(parsed.subject || '');
  const normSection = (parsed.section || '').trim();
  const names = new Set();
  try {
    const allDetails = localStorage.getItem(CACHE_KEYS.journalDetails);
    if (!allDetails) return [];
    const details = JSON.parse(allDetails);
    for (const item of parsedItems) {
      const itemSubject = normalizeSubject(item.subject || '');
      if (itemSubject !== normSubject) continue;
      for (const c of item.classes || []) {
        const classLabel = c.classLabel || '';
        if (!classLabelMatches(normSection, classLabel)) continue;
        const key = `${c.journalId}-${c.subgroupId || ''}`;
        const d = details[key];
        if (d?.students) {
          for (const s of d.students) {
            if (s?.name) names.add(s.name);
          }
        }
      }
    }
  } catch (_) {}
  return [...names].sort((a, b) => (a || '').localeCompare(b || '', 'uk'));
}

function getStudentsFromJournalsForClass(className) {
  const normClass = normalizeClassLabel(className || '');
  const exactClass = (className || '').trim();
  const names = new Set();
  try {
    const journalRows = parsedItems.flatMap(({ classes }) =>
      (classes || []).map(c => ({ journalId: c.journalId, subgroupId: c.subgroupId, classLabel: c.classLabel }))
    );
    const allDetails = localStorage.getItem(CACHE_KEYS.journalDetails);
    if (!allDetails) return [];
    const details = JSON.parse(allDetails);
    const useExact = exactClass && journalRows.some(j => (j.classLabel || '').trim() === exactClass);
    for (const j of journalRows) {
      const match = useExact
        ? (j.classLabel || '').trim() === exactClass
        : normalizeClassLabel(j.classLabel || '') === normClass;
      if (!match) continue;
      const key = `${j.journalId}-${j.subgroupId || ''}`;
      const d = details[key];
      if (d?.students) {
        for (const s of d.students) {
          if (s?.name) names.add(s.name);
        }
      }
    }
  } catch (_) {}
  return [...names].sort((a, b) => (a || '').localeCompare(b || '', 'uk'));
}

function getStudentInOtherClass(studentName, currentOrgPath) {
  try {
    const all = localStorage.getItem(CACHE_KEYS.classStudents);
    if (!all) return null;
    const obj = JSON.parse(all);
    const targetKey = normalizeNameForMatch(studentName);
    for (const [path, students] of Object.entries(obj)) {
      if (path === currentOrgPath || !students) continue;
      const found = students.find(s => {
        const key = normalizeNameForMatch(`${(s.familyName || '').trim()} ${(s.givenName || '').trim()}`);
        return key === targetKey || normalizeNameForMatch(`${(s.givenName || '').trim()} ${(s.familyName || '').trim()}`) === targetKey;
      });
      if (found) return { ...found, fromPath: path };
    }
  } catch (_) {}
  return null;
}

function getStudentMatchSets(googleStudents, journalStudents) {
  const googleKeys = new Set();
  (googleStudents || []).forEach(s => {
    const fn = (s.familyName || '').trim();
    const gn = (s.givenName || '').trim();
    if (fn || gn) {
      googleKeys.add(normalizeNameForMatch(`${fn} ${gn}`));
      googleKeys.add(normalizeNameForMatch(`${gn} ${fn}`));
    }
  });
  const journalKeys = new Set();
  const journalKeysCorrect = new Set();
  (journalStudents || []).forEach(name => {
    const n = normalizeNameForMatch(name);
    journalKeys.add(n);
    journalKeysCorrect.add(n);
    const parts = (name || '').trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) journalKeys.add(normalizeNameForMatch([...parts.slice(1), parts[0]].join(' ')));
  });
  return { googleKeys, journalKeys, journalKeysCorrect };
}

function getClassesFromJournalsCount() {
  const labels = new Set(parsedItems.flatMap(i => (i.classes || []).map(c => normalizeClassLabel(c.classLabel)).filter(Boolean)));
  return labels.size;
}

function getMismatchClasses() {
  const adminMismatches = [];
  const classroomMismatches = [];
  for (const c of adminClasses) {
    const name = normalizeClassLabel(c.name || '') || c.name || '';
    const path = c.orgUnitPath || '';
    if (!path) continue;
    const students = getClassStudentsCache(path);
    if (!students) continue;
    const journalStudents = getStudentsFromJournalsForClass(name);
    if (!journalStudents.length) continue;
    const { googleKeys, journalKeys } = getStudentMatchSets(students, journalStudents);
    const missingInJournal = (students || []).some(s => {
      const fn = (s.familyName || '').trim();
      const gn = (s.givenName || '').trim();
      const key1 = normalizeNameForMatch(`${fn} ${gn}`);
      const key2 = normalizeNameForMatch(`${gn} ${fn}`);
      return journalKeys.size > 0 && !journalKeys.has(key1) && !journalKeys.has(key2);
    });
    const missingInGoogle = (journalStudents || []).some(n => journalKeys.size > 0 && googleKeys.size > 0 && !googleKeys.has(normalizeNameForMatch(n)));
    if (missingInJournal || missingInGoogle) {
      adminMismatches.push({ name, path, missingInJournal, missingInGoogle });
    }
  }
  for (const course of classroomCourses) {
    const courseId = course.id || course.courseId;
    const name = course.name || '';
    const section = course.section || '';
    if (!name && !section) continue;
    const roster = courseRosterCache[courseId] || getCourseRosterCache(courseId);
    const students = roster?.students || [];
    const journalStudents = getStudentsFromJournalsForCourse(name, section);
    if (!journalStudents.length) continue;
    const { googleKeys, journalKeys } = getStudentMatchSetsFromNames(students, journalStudents);
    const missingInJournal = (students || []).some(s => {
      const fn = (s.familyName || parseStudentName(s.name || '').familyName || '').trim();
      const gn = (s.givenName || parseStudentName(s.name || '').givenName || '').trim();
      const key1 = normalizeNameForMatch(`${fn} ${gn}`);
      const key2 = normalizeNameForMatch(`${gn} ${fn}`);
      return journalKeys.size > 0 && !journalKeys.has(key1) && !journalKeys.has(key2);
    });
    const missingInGoogle = (journalStudents || []).some(n => journalKeys.size > 0 && googleKeys.size > 0 && !googleKeys.has(normalizeNameForMatch(n)));
    if (missingInJournal || missingInGoogle) {
      const displayName = (section && name) ? `${section} — ${name}` : (section || name);
      classroomMismatches.push({ name: displayName, courseId, missingInJournal, missingInGoogle });
    }
  }
  return { adminMismatches, classroomMismatches };
}

function renderDashboardMismatches() {
  const el = document.getElementById('dashboardMismatches');
  if (!el) return;
  const isLoading = courseRosterLoading.size > 0 || classStudentsLoading.size > 0;
  if (isLoading) {
    el.innerHTML = '<p class="status loading" style="margin:0;">Завантаження...</p>';
    return;
  }
  const { adminMismatches, classroomMismatches } = getMismatchClasses();
  const total = adminMismatches.length + classroomMismatches.length;
  if (!total) {
    const coursesWithoutRoster = classroomCourses.filter(c => {
      const id = c.id || c.courseId;
      return id && !getCourseRosterCache(id);
    });
    if (coursesWithoutRoster.length && tokens) {
      el.innerHTML = '<p class="status loading" style="margin:0;">Завантаження даних курсів...</p>';
      preloadCourseRosters();
      return;
    }
    let hasData = false;
    try {
      const classStudents = JSON.parse(localStorage.getItem(CACHE_KEYS.classStudents) || '{}');
      const courseRosters = JSON.parse(localStorage.getItem(CACHE_KEYS.courseRosters) || '{}');
      hasData = (adminClasses.length && Object.keys(classStudents).length) || (classroomCourses.length && Object.keys(courseRosters).length);
    } catch (_) {}
    el.innerHTML = hasData
      ? '<p class="status success" style="margin:0;">Невідповідностей не знайдено</p>'
      : '<p class="status" style="margin:0;color:var(--text-muted);">Завантажте класи та курси для перевірки</p>';
    updateManagePanelState();
    return;
  }
  let html = '<div class="mismatch-grid">';
  if (adminMismatches.length) {
    html += '<div class="mismatch-section"><div class="mismatch-label">Google Admin ↔ Журнал</div><ul class="mismatch-list">';
    adminMismatches.forEach(({ name, path }) => {
      html += `<li class="mismatch-item" data-panel="classes" data-path="${escapeHtml(path)}">${escapeHtml(name)}</li>`;
    });
    html += '</ul></div>';
  }
  if (classroomMismatches.length) {
    html += '<div class="mismatch-section"><div class="mismatch-label">Classroom ↔ Журнал</div><ul class="mismatch-list">';
    classroomMismatches.forEach(({ name, courseId }) => {
      html += `<li class="mismatch-item" data-panel="classroom" data-course-id="${escapeHtml(courseId)}">${escapeHtml(name)}</li>`;
    });
    html += '</ul></div>';
  }
  html += '</div>';
  el.innerHTML = html;
  markCourseRowsMismatches();
  updateManagePanelState();
  el.querySelectorAll('.mismatch-item').forEach(item => {
    item.onclick = () => {
      const panel = item.dataset.panel;
      showPanel(panel);
      if (panel === 'classes') {
        const path = item.dataset.path || '';
        const row = [...document.querySelectorAll('.class-row')].find(r => r.dataset.path === path);
        if (row) {
          if (!row.classList.contains('expanded')) row.click();
          setTimeout(() => row.scrollIntoView({ behavior: 'smooth', block: 'center' }), 100);
        }
      } else if (panel === 'classroom') {
        const cid = item.dataset.courseId || '';
        const row = [...document.querySelectorAll('.course-row')].find(r => (r.dataset.id || '') === cid);
        if (row) {
          if (!row.classList.contains('expanded')) row.click();
          setTimeout(() => row.scrollIntoView({ behavior: 'smooth', block: 'center' }), 100);
        }
      }
    };
  });
}

function updateStats() {
  const s1 = document.getElementById('statSubjects');
  const s2g = document.getElementById('statTeachersGoogle');
  const s2j = document.getElementById('statTeachersJournals');
  const s3g = document.getElementById('statClassesGoogle');
  const s3j = document.getElementById('statClassesJournals');
  const s4g = document.getElementById('statStudentsGoogle');
  const s4j = document.getElementById('statStudentsJournals');
  const jStats = getJournalStats();
  const journalDetailsLoaded = areAllJournalDetailsLoaded();
  if (s1) s1.textContent = parsedItems.length;
  if (s2g) s2g.textContent = teachers.length;
  if (s2j) s2j.textContent = journalDetailsLoaded ? (jStats.teachersCount ?? 0) : 'Завантаження...';
  if (s3g) s3g.textContent = adminClassesCount;
  if (s3j) s3j.textContent = getClassesFromJournalsCount();
  if (s4g) s4g.textContent = adminStudentsCount;
  if (s4j) s4j.textContent = journalDetailsLoaded ? (jStats.studentsCount ?? 0) : 'Завантаження...';
  renderDashboardMismatches();
}

function renderJournalsTable() {
  const table = document.getElementById('parseTable');
  if (!table || !parsedItems.length) return;
  const journalRows = parsedItems.flatMap(({ subject, classes }) =>
    classes.map(c => ({ subject, journalId: c.journalId, subgroupId: c.subgroupId, classLabel: c.classLabel }))
  );
  const sorted = sortJournalRows(journalRows);
  table.style.display = 'block';
  table.innerHTML = '<table><thead><tr><th>Предмет</th><th>Клас</th></tr></thead><tbody>' +
    sorted.map((j, i) => `<tr class="journal-row" data-idx="${i}" data-journal-id="${escapeHtml(j.journalId)}" data-subgroup-id="${escapeHtml(j.subgroupId || '')}"><td>${escapeHtml(j.subject)}</td><td>${escapeHtml(j.classLabel)}</td></tr>`).join('') +
    '</tbody></table>';
  table.querySelectorAll('.journal-row').forEach(row => row.onclick = () => toggleJournalDetails(row));
  document.getElementById('cardTeachers').style.display = 'block';
  document.getElementById('cardClasses').style.display = 'block';
  const cardJournalTeachers = document.getElementById('cardTeachersJournals');
  if (cardJournalTeachers) cardJournalTeachers.style.display = 'block';
  renderJournalTeachersTable();
}

function getTeachersFromJournals() {
  const names = new Set();
  try {
    const all = localStorage.getItem(CACHE_KEYS.journalDetails);
    if (all) {
      const obj = JSON.parse(all);
      Object.values(obj).forEach(d => {
        if (d?.teacher?.name) names.add(d.teacher.name);
      });
    }
  } catch (_) {}
  return [...names].sort((a, b) => (a || '').localeCompare(b || '', 'uk'));
}

function normalizeNameForMatch(s) {
  return (s || '')
    .replace(/[\u0027\u2019\u02BC\u0060\u00B4\u2032]/g, "'")
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function getTeacherMatchSets() {
  const googleKeys = new Set();
  teachers.forEach(t => {
    const fn = (t.familyName || '').trim();
    const gn = (t.givenName || '').trim();
    if (fn || gn) {
      googleKeys.add(normalizeNameForMatch(`${fn} ${gn}`));
      googleKeys.add(normalizeNameForMatch(`${gn} ${fn}`));
    }
  });
  const journalKeys = new Set();
  const journalKeysCorrect = new Set();
  getTeachersFromJournals().forEach(name => {
    const n = normalizeNameForMatch(name);
    journalKeys.add(n);
    journalKeysCorrect.add(n);
    const parts = (name || '').trim().split(/\s+/).filter(Boolean);
    if (parts.length >= 2) journalKeys.add(normalizeNameForMatch([...parts.slice(1), parts[0]].join(' ')));
  });
  return { googleKeys, journalKeys, journalKeysCorrect };
}

function renderTeachersTable() {
  const listEl = document.getElementById('adminTeachersList');
  if (!listEl) return;
  if (!teachers.length) {
    listEl.style.display = 'none';
    listEl.innerHTML = '';
    return;
  }
  const { journalKeys, journalKeysCorrect } = getTeacherMatchSets();
  const journalsReady = areAllJournalDetailsLoaded();
  const toSwap = [];
  const rows = teachers.map((t, i) => {
    const fn = (t.familyName || '').trim();
    const gn = (t.givenName || '').trim();
    const key1 = normalizeNameForMatch(`${fn} ${gn}`);
    const key2 = normalizeNameForMatch(`${gn} ${fn}`);
    const missing = journalsReady && !journalKeys.has(key1) && !journalKeys.has(key2) && journalKeys.size > 0;
    const needsSwap = !missing && journalsReady && (journalKeys.has(key1) || journalKeys.has(key2)) && !journalKeysCorrect.has(key1) && journalKeysCorrect.size > 0;
    if (needsSwap) toSwap.push({ email: t.email, familyName: t.familyName, givenName: t.givenName });
    const cls = missing ? ' class="teacher-missing"' : (needsSwap ? ' class="name-swapped"' : '');
    const removeBtn = missing ? `<button type="button" class="btn-remove-teacher btn-secondary" data-email="${escapeHtml(t.email)}" title="Перемістити в Вибувші">×</button>` : '';
    const editBtn = `<button type="button" class="btn-edit-user btn-secondary" data-email="${escapeHtml(t.email)}" data-family="${escapeHtml(t.familyName || '')}" data-given="${escapeHtml(t.givenName || '')}" data-type="teacher" title="Редагувати">✏️</button>`;
    const swapBtn = needsSwap ? `<button type="button" class="btn-swap-names btn-secondary" data-email="${escapeHtml(t.email)}" data-family="${escapeHtml(t.familyName || '')}" data-given="${escapeHtml(t.givenName || '')}" data-type="teacher" title="Поміняти місцями ім'я та прізвище">⇄</button>` : '';
    return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(t.familyName)}</td><td>${escapeHtml(t.givenName)}</td><td>${escapeHtml(t.email)}</td><td class="btn-cell"><div class="btn-cell-inner">${editBtn}${swapBtn}${removeBtn}</div></td></tr>`;
  });
  const lastTh = toSwap.length > 3 ? `<th style="text-align:right;"><button type="button" class="btn-mass-swap btn-secondary btn-sm" id="btnMassSwapTeachers" title="Поміняти місцями ім'я та прізвище у всіх">⇄ Усіх (${toSwap.length})</button></th>` : '<th></th>';
  listEl.style.display = 'block';
  listEl.innerHTML = '<table><thead><tr><th>#</th><th>Прізвище</th><th>Ім\'я</th><th>Email</th>' + lastTh + '</tr></thead><tbody>' + rows.join('') + '</tbody></table>';
  listEl.querySelectorAll('.btn-edit-user').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      openEditUserModal(btn.dataset.email, btn.dataset.family, btn.dataset.given, 'teacher');
    };
  });
  listEl.querySelectorAll('.btn-swap-names').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      swapUserNames(btn.dataset.email, btn.dataset.family, btn.dataset.given, 'teacher');
    };
  });
  listEl.querySelectorAll('.btn-remove-teacher').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      moveTeacherToGraduated(btn.dataset.email);
    };
  });
  const btnMass = listEl.querySelector('#btnMassSwapTeachers');
  if (btnMass) btnMass.onclick = () => swapUserNamesBulk(toSwap, 'teacher');
}

function closeAddTeacherToCourseModal() {
  const modal = document.getElementById('modalAddTeacherToCourse');
  if (modal) modal.style.display = 'none';
}

async function openAddTeacherToCourseModal(courseId, courseName, courseSection, courseRow) {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  const modal = document.getElementById('modalAddTeacherToCourse');
  const title = document.getElementById('modalAddTeacherToCourseTitle');
  const content = document.getElementById('modalAddTeacherToCourseContent');
  const closeBtn = document.getElementById('btnModalAddTeacherToCourseClose');
  const sectionPart = (courseSection || '').trim();
  const namePart = (courseName || '').trim();
  const displayTitle = sectionPart && namePart && namePart !== sectionPart ? sectionPart + ' — ' + namePart : (sectionPart || namePart || '');
  title.textContent = 'Додати викладача до курсу: ' + displayTitle;
  content.innerHTML = '<p class="status loading" style="margin:0;">Завантаження...</p>';
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeAddTeacherToCourseModal(); };
  if (closeBtn) closeBtn.onclick = closeAddTeacherToCourseModal;
  try {
    let adminTeachers = teachers;
    if (!adminTeachers.length) {
      const r = await fetch(`${API}/teachers/admin`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, ouName: 'Вчителі' })
      });
      const data = await r.json().catch(() => ({}));
      if (!data.error) {
        adminTeachers = (data.teachers || []).slice().sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
        teachers = adminTeachers;
      }
    }
    const courseTeacherEmails = new Set((courseRosterCache[courseId]?.teachers || []).map(t => (t.email || '').toLowerCase()).filter(Boolean));
    const available = adminTeachers.filter(t => {
      const em = (t.email || '').toLowerCase();
      return em && !courseTeacherEmails.has(em);
    });
    if (!available.length) {
      content.innerHTML = '<p style="color:var(--text-muted);margin:0;">Усі викладачі з Google Admin вже додані до курсу або список порожній.</p>';
      return;
    }
    content.innerHTML = '<table style="font-size:0.9rem;"><thead><tr><th>ПІБ</th><th>Email</th><th></th></tr></thead><tbody>' +
      available.map(t => {
        const name = [t.familyName, t.givenName].filter(Boolean).join(' ') || t.email || '';
        return `<tr><td>${escapeHtml(name)}</td><td>${escapeHtml(t.email || '')}</td><td><button type="button" class="btn-add-teacher-to-course-row btn-secondary btn-sm" data-email="${escapeHtml(t.email)}" data-family="${escapeHtml(t.familyName || '')}" data-given="${escapeHtml(t.givenName || '')}">Додати</button></td></tr>`;
      }).join('') +
      '</tbody></table>';
    content.querySelectorAll('.btn-add-teacher-to-course-row').forEach(btn => {
      btn.onclick = async () => {
        const email = btn.dataset.email;
        if (!email) return;
        btn.disabled = true;
        btn.textContent = '...';
        try {
          const r = await fetch(`${API}/classroom/course/add-teacher`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ tokens, courseId, email })
          });
          const text = await r.text();
          let data;
          try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
          if (data.error) throw new Error(data.error);
          closeAddTeacherToCourseModal();
          if (courseRosterCache[courseId]) {
            const newTeacher = { email, name: [btn.dataset.family, btn.dataset.given].filter(Boolean).join(' ') || email };
            courseRosterCache[courseId].teachers = (courseRosterCache[courseId].teachers || []).concat(newTeacher);
            courseRosterCache[courseId].teachers.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'uk'));
            setCourseRosterCache(courseId, courseRosterCache[courseId]);
          }
          const detailRow = courseRow.nextElementSibling;
          if (detailRow?.classList?.contains('course-detail-row')) detailRow.remove();
          renderCourseRosterRow(courseRow, courseRow.dataset.name, courseRow.dataset.section || '', courseRosterCache[courseId]?.teachers || [], courseRosterCache[courseId]?.students || []);
          renderDashboardMismatches();
          showAlertModal('Викладача додано до курсу.');
        } catch (e) {
          showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
        } finally {
          btn.disabled = false;
          btn.textContent = 'Додати';
        }
      };
    });
  } catch (e) {
    content.innerHTML = '<p class="status error" style="margin:0;">' + escapeHtml(e.message || e.name) + '</p>';
  }
}

async function openAddTeacherToMassCoursesModal(courseIds, courseRows) {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  if (!courseIds?.length) return;
  const modal = document.getElementById('modalAddTeacherToCourse');
  const title = document.getElementById('modalAddTeacherToCourseTitle');
  const content = document.getElementById('modalAddTeacherToCourseContent');
  const closeBtn = document.getElementById('btnModalAddTeacherToCourseClose');
  title.textContent = `Додати викладача до ${courseIds.length} обраних курсів`;
  content.innerHTML = '<p class="status loading" style="margin:0;">Завантаження...</p>';
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeAddTeacherToCourseModal(); };
  if (closeBtn) closeBtn.onclick = closeAddTeacherToCourseModal;
  try {
    let adminTeachers = teachers;
    if (!adminTeachers.length) {
      const r = await fetch(`${API}/teachers/admin`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, ouName: 'Вчителі' })
      });
      const data = await r.json().catch(() => ({}));
      if (!data.error) {
        adminTeachers = (data.teachers || []).slice().sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
        teachers = adminTeachers;
      }
    }
    const available = adminTeachers.filter(t => (t.email || '').trim());
    if (!available.length) {
      content.innerHTML = '<p style="color:var(--text-muted);margin:0;">Список викладачів порожній.</p>';
      return;
    }
    content.innerHTML = '<table style="font-size:0.9rem;"><thead><tr><th>ПІБ</th><th>Email</th><th></th></tr></thead><tbody>' +
      available.map(t => {
        const name = [t.familyName, t.givenName].filter(Boolean).join(' ') || t.email || '';
        return `<tr><td>${escapeHtml(name)}</td><td>${escapeHtml(t.email || '')}</td><td><button type="button" class="btn-add-teacher-mass-row btn-secondary btn-sm" data-email="${escapeHtml(t.email)}" data-family="${escapeHtml(t.familyName || '')}" data-given="${escapeHtml(t.givenName || '')}">Додати</button></td></tr>`;
      }).join('') +
      '</tbody></table>';
    content.querySelectorAll('.btn-add-teacher-mass-row').forEach(btn => {
      btn.onclick = async () => {
        const email = btn.dataset.email;
        if (!email) return;
        btn.disabled = true;
        btn.textContent = '...';
        try {
          let added = 0;
          for (let i = 0; i < courseIds.length; i++) {
            const cid = courseIds[i];
            const courseTeacherEmails = new Set((courseRosterCache[cid]?.teachers || []).map(t => (t.email || '').toLowerCase()).filter(Boolean));
            if (courseTeacherEmails.has(email.toLowerCase())) continue;
            const r = await fetch(`${API}/classroom/course/add-teacher`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ tokens, courseId: cid, email })
            });
            const data = await r.json().catch(() => ({}));
            if (!data.error) {
              added++;
              if (courseRosterCache[cid]) {
                const newTeacher = { email, name: [btn.dataset.family, btn.dataset.given].filter(Boolean).join(' ') || email };
                courseRosterCache[cid].teachers = (courseRosterCache[cid].teachers || []).concat(newTeacher);
                courseRosterCache[cid].teachers.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'uk'));
                setCourseRosterCache(cid, courseRosterCache[cid]);
              }
              const row = courseRows[i];
              if (row) {
                const detailRow = row.nextElementSibling;
                if (detailRow?.classList?.contains('course-detail-row')) detailRow.remove();
                renderCourseRosterRow(row, row.dataset.name, row.dataset.section || '', courseRosterCache[cid]?.teachers || [], courseRosterCache[cid]?.students || []);
              }
            }
          }
          closeAddTeacherToCourseModal();
          renderDashboardMismatches();
          showAlertModal(`Викладача додано до ${added} курсів.`);
        } catch (e) {
          showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
        } finally {
          btn.disabled = false;
          btn.textContent = 'Додати';
        }
      };
    });
  } catch (e) {
    content.innerHTML = '<p class="status error" style="margin:0;">' + escapeHtml(e.message || e.name) + '</p>';
  }
}

async function openFindStudentModal(studentName, className, orgUnitPath) {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  const modal = document.getElementById('modalFindStudent');
  const title = document.getElementById('modalFindStudentTitle');
  const content = document.getElementById('modalFindStudentContent');
  const actions = document.getElementById('modalFindStudentActions');
  title.textContent = `Пошук: ${studentName}`;
  content.innerHTML = '<p class="status loading" style="margin:0;">Пошук...</p>';
  actions.innerHTML = '<button class="btn-secondary" id="btnModalFindStudentClose">Закрити</button>';
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeFindStudentModal(); };
  document.getElementById('btnModalFindStudentClose').onclick = closeFindStudentModal;
  try {
    const parts = (studentName || '').trim().split(/\s+/).filter(Boolean);
    const searchTerms = [...new Set(parts)];
    const allUsers = [];
    for (const term of searchTerms) {
      const r = await fetch(`${API}/users/search`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, searchTerm: term })
      });
      const text = await r.text();
      let data;
      try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
      if (data.error) throw new Error(data.error);
      (data.users || []).forEach(u => {
        if (!allUsers.find(x => x.email === u.email)) allUsers.push(u);
      });
    }
    const targetKey = normalizeNameForMatch(studentName);
    const targetKeySwapped = parts.length >= 2 ? normalizeNameForMatch(parts.slice(1).concat(parts[0]).join(' ')) : targetKey;
    const matches = allUsers.filter(u => {
      const key1 = normalizeNameForMatch(`${(u.familyName || '').trim()} ${(u.givenName || '').trim()}`);
      const key2 = normalizeNameForMatch(`${(u.givenName || '').trim()} ${(u.familyName || '').trim()}`);
      return key1 === targetKey || key2 === targetKey || key1 === targetKeySwapped || key2 === targetKeySwapped;
    });
    if (matches.length) {
      const targetPath = (orgUnitPath || '').replace(/\/$/, '');
      content.innerHTML = '<table style="font-size:0.9rem;"><thead><tr><th>ПІБ</th><th>Email</th><th>Підрозділ</th><th></th></tr></thead><tbody>' +
        matches.map(u => {
          const userPath = (u.orgUnitPath || '').replace(/\/$/, '');
          const inOtherOu = targetPath && userPath !== targetPath;
          const moveBtn = inOtherOu ? `<button type="button" class="btn-move-student" data-email="${escapeHtml(u.email)}" data-path="${escapeHtml(orgUnitPath)}">Перемістити в цей клас</button>` : '';
          return `<tr><td>${escapeHtml((u.familyName || '') + ' ' + (u.givenName || ''))}</td><td>${escapeHtml(u.email)}</td><td style="font-size:0.8rem;color:var(--text-muted)">${escapeHtml(u.orgUnitPath || '')}</td><td>${moveBtn}</td></tr>`;
        }).join('') +
        '</tbody></table>';
      content.querySelectorAll('.btn-move-student').forEach(btn => {
        btn.onclick = async () => {
          const email = btn.dataset.email;
          const path = btn.dataset.path;
          if (!(await showConfirmModal(`Перемістити ${email} в цей клас?`))) return;
          btn.disabled = true;
          try {
            const r = await fetch(`${API}/users/move`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ tokens, email, orgUnitPath: path })
            });
            const text = await r.text();
            let data;
            try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
            if (data.error) throw new Error(data.error);
            closeFindStudentModal();
            const cached = getClassStudentsCache(path);
            if (cached) {
              const moved = matches.find(m => m.email === email);
              if (moved) {
                cached.push({ familyName: moved.familyName, givenName: moved.givenName, email: moved.email });
                cached.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
                setClassStudentsCache(path, cached);
              }
            }
            const classRow = document.querySelector(`.class-row[data-path="${path}"]`);
            if (classRow) {
              const detailRow = classRow.nextElementSibling;
              if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
              renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
            }
            showAlertModal('Користувача переміщено.');
          } catch (e) {
            showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
          } finally {
            btn.disabled = false;
          }
        };
      });
    } else {
      const { familyName, givenName } = parseStudentName(studentName);
      const email = generateTeacherEmail(familyName, givenName);
      const password = generateTeacherPassword();
      content.innerHTML = '<p style="color:var(--text-muted);margin-bottom:1rem;">Збігів не знайдено.</p>' +
        '<div class="modal-credentials"><div class="modal-credentials-warn">⚠ Створити учня в класі ' + escapeHtml(className || '') + '? Збережіть дані!</div>' +
        '<div class="modal-row"><label>ПІБ</label><input type="text" class="readonly" readonly value="' + escapeHtml(studentName) + '"></div>' +
        '<div class="modal-row"><label>Email</label><input type="text" id="findStudentEmail" class="readonly" readonly value="' + escapeHtml(email) + '"></div>' +
        '<div class="modal-row"><label>Пароль</label><input type="text" id="findStudentPassword" class="readonly" readonly value="' + escapeHtml(password) + '"></div></div>';
      actions.innerHTML = '<button class="btn-secondary" id="btnModalFindStudentClose">Закрити</button><button class="btn-primary" id="btnModalFindStudentCreate">Створити</button>';
      document.getElementById('btnModalFindStudentClose').onclick = closeFindStudentModal;
      document.getElementById('btnModalFindStudentCreate').onclick = async () => {
        const btn = document.getElementById('btnModalFindStudentCreate');
        btn.disabled = true;
        btn.textContent = 'Створення...';
        try {
          const createR = await fetch(`${API}/users/create`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              tokens,
              givenName,
              familyName,
              primaryEmail: document.getElementById('findStudentEmail').value,
              password: document.getElementById('findStudentPassword').value,
              orgUnitPath
            })
          });
          const createText = await createR.text();
          let createData;
          try { createData = JSON.parse(createText); } catch (_) { throw new Error('Помилка відповіді'); }
          if (createData.error) throw new Error(createData.error);
          const newEmail = document.getElementById('findStudentEmail').value;
          const newPassword = document.getElementById('findStudentPassword').value;
          closeFindStudentModal();
          const path = orgUnitPath;
          if (path) {
            const cached = getClassStudentsCache(path);
            if (cached) {
              cached.push({ familyName, givenName, email: newEmail });
              cached.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
              setClassStudentsCache(path, cached);
            }
          }
          const classRow = document.querySelector(`.class-row[data-path="${path}"]`);
          if (classRow) {
            const detailRow = classRow.nextElementSibling;
            if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
            renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
          }
          showAlertModal('Учня створено.\nEmail: ' + newEmail + '\nПароль: ' + newPassword + '\n\nЗбережіть пароль!');
        } catch (e) {
          showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
        } finally {
          btn.disabled = false;
          btn.textContent = 'Створити';
        }
      };
    }
  } catch (e) {
    content.innerHTML = '<p class="status error" style="margin:0;">' + escapeHtml(e.message || e.name) + '</p>';
  }
}

function closeFindStudentModal() {
  document.getElementById('modalFindStudent').style.display = 'none';
}

async function openFindStudentForClassroomModal(studentName, courseName, courseId, courseRow) {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  const modal = document.getElementById('modalFindStudent');
  const title = document.getElementById('modalFindStudentTitle');
  const content = document.getElementById('modalFindStudentContent');
  const actions = document.getElementById('modalFindStudentActions');
  title.textContent = `Пошук: ${studentName}`;
  content.innerHTML = '<p class="status loading" style="margin:0;">Пошук...</p>';
  actions.innerHTML = '<button class="btn-secondary" id="btnModalFindStudentClose">Закрити</button>';
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeFindStudentModal(); };
  document.getElementById('btnModalFindStudentClose').onclick = closeFindStudentModal;
  try {
    const parts = (studentName || '').trim().split(/\s+/).filter(Boolean);
    const searchTerms = [...new Set(parts)];
    const allUsers = [];
    for (const term of searchTerms) {
      const r = await fetch(`${API}/users/search`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, searchTerm: term })
      });
      const text = await r.text();
      let data;
      try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
      if (data.error) throw new Error(data.error);
      (data.users || []).forEach(u => {
        if (!allUsers.find(x => x.email === u.email)) allUsers.push(u);
      });
    }
    const targetKey = normalizeNameForMatch(studentName);
    const targetKeySwapped = parts.length >= 2 ? normalizeNameForMatch(parts.slice(1).concat(parts[0]).join(' ')) : targetKey;
    const matches = allUsers.filter(u => {
      const key1 = normalizeNameForMatch(`${(u.familyName || '').trim()} ${(u.givenName || '').trim()}`);
      const key2 = normalizeNameForMatch(`${(u.givenName || '').trim()} ${(u.familyName || '').trim()}`);
      return key1 === targetKey || key2 === targetKey || key1 === targetKeySwapped || key2 === targetKeySwapped;
    });
    if (matches.length) {
      content.innerHTML = '<table style="font-size:0.9rem;"><thead><tr><th>ПІБ</th><th>Email</th><th>Підрозділ</th><th></th></tr></thead><tbody>' +
        matches.map(u => {
          const joinBtn = `<button type="button" class="btn-join-course btn-secondary" data-email="${escapeHtml(u.email)}" data-family="${escapeHtml(u.familyName || '')}" data-given="${escapeHtml(u.givenName || '')}">Приєднати до курсу</button>`;
          return `<tr><td>${escapeHtml((u.familyName || '') + ' ' + (u.givenName || ''))}</td><td>${escapeHtml(u.email)}</td><td style="font-size:0.8rem;color:var(--text-muted)">${escapeHtml(u.orgUnitPath || '')}</td><td>${joinBtn}</td></tr>`;
        }).join('') +
        '</tbody></table>';
      content.querySelectorAll('.btn-join-course').forEach(btn => {
        btn.onclick = async () => {
          const email = btn.dataset.email;
          if (!(await showConfirmModal(`Приєднати ${email} до курсу ${courseName || ''}?`))) return;
          btn.disabled = true;
          try {
            const r = await fetch(`${API}/classroom/course/add-student`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ tokens, courseId, email })
            });
            const text = await r.text();
            let data;
            try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
            if (data.error) throw new Error(data.error);
            closeFindStudentModal();
            if (courseRosterCache[courseId]) {
              const newStudent = { email, familyName: btn.dataset.family, givenName: btn.dataset.given };
              courseRosterCache[courseId].students.push(newStudent);
              courseRosterCache[courseId].students.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
              setCourseRosterCache(courseId, courseRosterCache[courseId]);
            }
            if (courseRow) {
              const detailRow = courseRow.nextElementSibling;
              if (detailRow?.classList?.contains('course-detail-row')) detailRow.remove();
              renderCourseRosterRow(courseRow, courseRow.dataset.name, courseRow.dataset.section || '', courseRosterCache[courseId]?.teachers || [], courseRosterCache[courseId]?.students || []);
            }
            renderDashboardMismatches();
            showAlertModal('Учня приєднано до курсу.');
          } catch (e) {
            showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
          } finally {
            btn.disabled = false;
          }
        };
      });
    } else {
      content.innerHTML = '<p style="color:var(--text-muted);margin:0;">Збігів не знайдено.</p>';
    }
  } catch (e) {
    content.innerHTML = '<p class="status error" style="margin:0;">' + escapeHtml(e.message || e.name) + '</p>';
  }
}

function openEditUserModal(email, familyName, givenName, type, path, classRow) {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  const modal = document.getElementById('modalEditUser');
  const fn = document.getElementById('editUserFamilyName');
  const gn = document.getElementById('editUserGivenName');
  const pwd = document.getElementById('editUserPassword');
  fn.value = familyName || '';
  gn.value = givenName || '';
  pwd.value = '';
  modal.dataset.editEmail = email;
  modal.dataset.editType = type;
  modal.dataset.editPath = path || '';
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeEditUserModal(); };
  document.getElementById('btnModalEditUserCancel').onclick = closeEditUserModal;
  document.getElementById('btnModalEditUserSubmit').onclick = () => submitEditUser(modal);
}

function closeEditUserModal() {
  document.getElementById('modalEditUser').style.display = 'none';
}

async function swapUserNamesBulk(items, type, path) {
  if (!loadTokens() || !items?.length) return;
  const count = items.length;
  if (!(await showConfirmModal(`Поміняти місцями ім'я та прізвище у ${count} користувачів?`))) return;
  let done = 0;
  for (const it of items) {
    try {
      const r = await fetch(`${API}/users/update`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tokens, email: it.email, familyName: it.givenName, givenName: it.familyName })
      });
      const text = await r.text();
      let data;
      try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
      if (data.error) throw new Error(data.error);
      if (type === 'teacher') {
        const t = teachers.find(x => x.email === it.email);
        if (t) { t.familyName = it.givenName; t.givenName = it.familyName; }
      } else if (type === 'student' && path) {
        const cached = getClassStudentsCache(path);
        if (cached) {
          const s = cached.find(x => x.email === it.email);
          if (s) { s.familyName = it.givenName; s.givenName = it.familyName; }
          cached.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
          setClassStudentsCache(path, cached);
        }
      }
      done++;
    } catch (e) {
      showAlertModal('Помилка для ' + it.email + ': ' + (e.message || e.name), 'Помилка', true);
      break;
    }
  }
  if (type === 'teacher') {
    teachers.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
    renderTeachersTable();
    renderJournalTeachersTable();
  } else if (type === 'student' && path) {
    const classRow = document.querySelector(`.class-row[data-path="${path}"]`);
    if (classRow) {
      const detailRow = classRow.nextElementSibling;
      if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
      renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
    }
  }
  updateStats();
  saveToCache();
  showAlertModal(`Оновлено ${done} з ${count}.`);
}

async function swapUserNames(email, familyName, givenName, type, path) {
  if (!loadTokens() || !email) return;
  const fn = (familyName || '').trim();
  const gn = (givenName || '').trim();
  if (!fn && !gn) return;
  if (!(await showConfirmModal(`Поміняти місцями прізвище та ім'я для ${email}?`))) return;
  try {
    const r = await fetch(`${API}/users/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, email, familyName: gn, givenName: fn })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    if (type === 'teacher') {
      const t = teachers.find(x => x.email === email);
      if (t) { t.familyName = gn; t.givenName = fn; }
      teachers.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
      renderTeachersTable();
      renderJournalTeachersTable();
    } else if (type === 'student' && path) {
      const cached = getClassStudentsCache(path);
      if (cached) {
        const s = cached.find(x => x.email === email);
        if (s) { s.familyName = gn; s.givenName = fn; }
        cached.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
        setClassStudentsCache(path, cached);
      }
      const classRow = document.querySelector(`.class-row[data-path="${path}"]`);
      if (classRow) {
        const detailRow = classRow.nextElementSibling;
        if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
        renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
      }
    }
    updateStats();
    saveToCache();
    showAlertModal('Збережено.');
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
}

async function submitEditUser(modal) {
  const email = modal.dataset.editEmail;
  const type = modal.dataset.editType;
  const path = modal.dataset.editPath;
  const fn = document.getElementById('editUserFamilyName').value.trim();
  const gn = document.getElementById('editUserGivenName').value.trim();
  const pwd = document.getElementById('editUserPassword').value;
  if (!fn || !gn) { showAlertModal('Введіть прізвище та ім\'я', 'Увага'); return; }
  try {
    const r = await fetch(`${API}/users/update`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, email, familyName: fn, givenName: gn, password: pwd || undefined })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    closeEditUserModal();
    if (type === 'teacher') {
      const t = teachers.find(x => x.email === email);
      if (t) { t.familyName = fn; t.givenName = gn; }
      teachers.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
      renderTeachersTable();
      renderJournalTeachersTable();
    } else if (type === 'student' && path) {
      const cached = getClassStudentsCache(path);
      if (cached) {
        const s = cached.find(x => x.email === email);
        if (s) { s.familyName = fn; s.givenName = gn; }
        cached.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
        setClassStudentsCache(path, cached);
      }
      const classRow = document.querySelector(`.class-row[data-path="${path}"]`);
      if (classRow) {
        const detailRow = classRow.nextElementSibling;
        if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
        renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
      }
    }
    updateStats();
    saveToCache();
    showAlertModal('Збережено.');
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
}

async function moveTeacherToGraduated(email) {
  if (!loadTokens() || !email) return;
  if (!(await showConfirmModal(`Перемістити вчителя ${email} в організаційний підрозділ «Вибувші»?`))) return;
  try {
    const r = await fetch(`${API}/teachers/move-to-graduated`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, email })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    teachers = teachers.filter(t => t.email !== email);
    renderTeachersTable();
    renderJournalTeachersTable();
    updateStats();
    saveToCache();
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
}

function renderJournalTeachersTable() {
  const listEl = document.getElementById('journalTeachersList');
  if (!listEl) return;
  if (!areAllJournalDetailsLoaded()) {
    listEl.innerHTML = '<p class="status loading" style="margin:0;">Завантаження...</p>';
    return;
  }
  const journalTeachers = getTeachersFromJournals();
  if (!journalTeachers.length) {
    listEl.innerHTML = '<p style="color:var(--text-muted);margin:0;">Немає даних</p>';
    return;
  }
  const { googleKeys } = getTeacherMatchSets();
  listEl.innerHTML = '<table><thead><tr><th>#</th><th>ПІБ</th></tr></thead><tbody>' +
    journalTeachers.map((name, i) => {
      const key = normalizeNameForMatch(name);
      const missing = !googleKeys.has(key) && googleKeys.size > 0;
      const cls = missing ? ' class="teacher-missing"' : '';
      return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(name)}</td></tr>`;
    }).join('') +
    '</tbody></table>';
}

function sortClassesByGrade(classes) {
  const parseClass = (s) => {
    const norm = normalizeClassLabel(s || '');
    const m = norm.match(/^(\d+)[-\s]*(.*)$/);
    return m ? { grade: parseInt(m[1], 10), letter: (m[2] || '').trim() } : { grade: 0, letter: norm };
  };
  return classes.slice().sort((a, b) => {
    const pa = parseClass(a.name);
    const pb = parseClass(b.name);
    if (pa.grade !== pb.grade) return pb.grade - pa.grade;
    return (pa.letter || '').localeCompare(pb.letter || '', 'uk');
  });
}

function renderClassesTable() {
  const listEl = document.getElementById('adminClassesList');
  if (!listEl || !adminClasses.length) return;
  listEl.style.display = 'block';
  const sorted = sortClassesByGrade(adminClasses);
  listEl.innerHTML = '<table><thead><tr><th>Клас</th></tr></thead><tbody>' +
    sorted.map((c, i) => {
      const displayName = normalizeClassLabel(c.name || '') || c.name || '';
      return `<tr class="class-row" data-path="${escapeHtml(c.orgUnitPath || '')}" data-name="${escapeHtml(displayName)}" data-idx="${i}"><td>${escapeHtml(displayName)}</td></tr>`;
    }).join('') +
    '</tbody></table>';
  listEl.querySelectorAll('.class-row').forEach(row => row.onclick = () => toggleClassStudents(row));
}

function sortCoursesBySection(courses) {
  const collator = new Intl.Collator('uk');
  const parseClass = (s) => {
    const norm = normalizeClassLabel(s || '');
    const m = norm.match(/(\d+)[-\s]*([А-Яа-яІіЄєЇїҐґA-Za-z]+)/);
    return m ? { grade: parseInt(m[1], 10), letter: (m[2] || '').trim() } : { grade: 0, letter: norm };
  };
  return courses.slice().sort((a, b) => {
    const sa = a.section || '';
    const sb = b.section || '';
    const pa = parseClass(sa);
    const pb = parseClass(sb);
    if (pa.grade !== pb.grade) return pb.grade - pa.grade;
    return collator.compare(sa, sb);
  });
}

function renderClassroomTable() {
  const listEl = document.getElementById('classroomCoursesList');
  const massEl = document.getElementById('classroomCoursesMassActions');
  if (!listEl || !classroomCourses.length) return;
  listEl.style.display = 'block';
  const sorted = sortCoursesBySection(classroomCourses);
  listEl.innerHTML = '<table><thead><tr><th style="width:2.5rem;"><input type="checkbox" id="courseSelectAll" title="Обрати всі"></th><th>Секція</th><th>Курс</th></tr></thead><tbody>' +
    sorted.map((c, i) => {
      const name = escapeHtml(c.name || '');
      const section = escapeHtml(c.section || '');
      const cid = escapeHtml(c.id || '');
      return `<tr class="course-row" data-id="${cid}" data-name="${name}" data-section="${escapeHtml(c.section || '')}" data-idx="${i}"><td class="course-check-cell"><input type="checkbox" class="course-checkbox" data-course-id="${cid}"></td><td>${section}</td><td>${name}</td></tr>`;
    }).join('') +
    '</tbody></table>';
  listEl.querySelectorAll('.course-row').forEach(row => {
    row.onclick = (e) => {
      if (e.target.closest('.course-check-cell')) return;
      toggleCourseRoster(row);
    };
  });
  const selectAll = document.getElementById('courseSelectAll');
  if (selectAll) {
    selectAll.checked = false;
    selectAll.indeterminate = false;
    selectAll.onclick = (e) => {
      e.stopPropagation();
      const checked = selectAll.checked;
      listEl.querySelectorAll('.course-checkbox').forEach(cb => { cb.checked = checked; });
      updateClassroomMassActions();
    };
  }
  listEl.querySelectorAll('.course-checkbox').forEach(cb => {
    cb.onclick = (e) => { e.stopPropagation(); updateClassroomMassActions(); };
  });
  updateClassroomMassActions();
  markCourseRowsMismatches();
}

function markCourseRowsMismatches() {
  const { classroomMismatches } = getMismatchClasses();
  const mismatchIds = new Set(classroomMismatches.map(m => m.courseId));
  document.querySelectorAll('.course-row').forEach(row => {
    const id = row.dataset.id || '';
    row.classList.toggle('has-mismatch', mismatchIds.has(id));
  });
}

function updateClassroomMassActions() {
  const listEl = document.getElementById('classroomCoursesList');
  const massEl = document.getElementById('classroomCoursesMassActions');
  const selectAll = document.getElementById('courseSelectAll');
  if (!listEl || !massEl) return;
  const checkboxes = listEl.querySelectorAll('.course-checkbox');
  const checked = [...checkboxes].filter(cb => cb.checked);
  massEl.style.display = checked.length ? 'flex' : 'none';
  if (selectAll && checkboxes.length) {
    selectAll.checked = checked.length === checkboxes.length;
    selectAll.indeterminate = checked.length > 0 && checked.length < checkboxes.length;
  }
}

function formatRosterEmail(v) {
  if (!v) return '—';
  if (/^\d{15,}$/.test(String(v))) return '—';
  return escapeHtml(v);
}

function getStudentMatchSetsFromNames(googleStudentsWithName, journalStudents) {
  const googleStudents = (googleStudentsWithName || []).map(s => ({
    familyName: s.familyName || parseStudentName(s.name || '').familyName,
    givenName: s.givenName || parseStudentName(s.name || '').givenName
  }));
  return getStudentMatchSets(googleStudents, journalStudents);
}

function renderCourseRosterRow(row, name, section, teachers, students) {
  const tr = document.createElement('tr');
  tr.className = 'course-detail-row';
  const courseId = row.dataset.id || '';
  const teachersRows = (teachers || []).map((t, i) => {
    const displayName = formatTeacherDisplayName(t.name || '');
    const removeBtn = `<button type="button" class="btn-remove-teacher-from-course btn-secondary" data-email="${escapeHtml(t.email)}" data-course-id="${escapeHtml(courseId)}" title="Видалити з курсу">×</button>`;
    return `<tr><td>${i + 1}</td><td>${escapeHtml(displayName)}</td><td>${formatRosterEmail(t.email)}</td><td class="btn-cell"><div class="btn-cell-inner">${removeBtn}</div></td></tr>`;
  }).join('');
  const journalStudents = getStudentsFromJournalsForCourse(name || '', section || '');
  const sortedStudents = (students || []).slice().sort((a, b) => {
    const fa = (a.familyName || a.name || '').trim();
    const fb = (b.familyName || b.name || '').trim();
    const cmp = (fa || '').localeCompare(fb || '', 'uk');
    if (cmp !== 0) return cmp;
    return (a.givenName || '').localeCompare(b.givenName || '', 'uk');
  });
  const { googleKeys, journalKeys, journalKeysCorrect } = getStudentMatchSetsFromNames(sortedStudents, journalStudents);
  const studentsRows = sortedStudents.map((s, i) => {
    const fn = s.familyName || parseStudentName(s.name || '').familyName;
    const gn = s.givenName || parseStudentName(s.name || '').givenName;
    const displayName = [fn, gn].filter(Boolean).join(' ') || s.name || '';
    const key1 = normalizeNameForMatch(`${fn} ${gn}`);
    const key2 = normalizeNameForMatch(`${gn} ${fn}`);
    const missing = !journalKeys.has(key1) && !journalKeys.has(key2) && journalKeys.size > 0;
    const needsSwap = !missing && (journalKeys.has(key1) || journalKeys.has(key2)) && !journalKeysCorrect.has(key1) && journalKeysCorrect.size > 0;
    const cls = missing ? ' class="teacher-missing"' : (needsSwap ? ' class="name-swapped"' : '');
    const removeBtn = missing ? `<button type="button" class="btn-remove-from-course btn-secondary" data-email="${escapeHtml(s.email)}" data-course-id="${escapeHtml(courseId)}" title="Видалити з курсу">×</button>` : '';
    return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(displayName)}</td><td>${formatRosterEmail(s.email)}</td><td class="btn-cell"><div class="btn-cell-inner">${removeBtn}</div></td></tr>`;
  }).join('');
  const journalRows = journalStudents.map((studentName, i) => {
    const key = normalizeNameForMatch(studentName);
    const missing = !googleKeys.has(key) && googleKeys.size > 0;
    const cls = missing ? ' class="teacher-missing"' : '';
    const btn = missing ? `<button type="button" class="btn-find-student btn-secondary" data-name="${escapeHtml(studentName)}" data-course="${escapeHtml(name)}" data-course-id="${escapeHtml(courseId)}" title="Знайти в Google Admin">🔍</button>` : '';
    return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(studentName)}</td><td class="btn-cell">${btn}</td></tr>`;
  }).join('');
  const addTeacherBtn = '<button type="button" class="btn-add-teacher-to-course btn-secondary btn-sm" data-course-id="' + escapeHtml(courseId) + '" data-course-name="' + escapeHtml(name) + '" data-course-section="' + escapeHtml(section || '') + '" title="Додати викладача">+ Додати</button>';
  const teachersTable = teachers?.length
    ? '<table class="class-students-table course-roster-table"><thead><tr><th>#</th><th>Прізвище та Ім\'я</th><th>Email</th><th></th></tr></thead><tbody>' + teachersRows + '</tbody></table>'
    : '<p style="color:var(--text-muted);margin:0;">Немає викладачів</p>';
  const teachersLabelHtml = '<div class="class-students-label" style="display:flex;align-items:center;gap:0.5rem;"><span>Викладачі</span>' + addTeacherBtn + '</div>';
  const studentsTable = students?.length ? '<table class="class-students-table course-roster-table"><thead><tr><th>#</th><th>Прізвище та Ім\'я</th><th>Email</th><th></th></tr></thead><tbody>' + studentsRows + '</tbody></table>' : '<p style="color:var(--text-muted);margin:0;">Немає учнів</p>';
  const journalTable = journalStudents.length ? '<table class="class-students-table course-roster-table"><thead><tr><th>#</th><th>Прізвище та Ім\'я</th><th></th></tr></thead><tbody>' + journalRows + '</tbody></table>' : (areAllJournalDetailsLoaded() ? '<p style="color:var(--text-muted);margin:0;">Немає даних</p>' : '<p class="status loading" style="margin:0;">Завантаження...</p>');
  tr.innerHTML = '<td colspan="2"><div class="class-students"><strong>' + escapeHtml(name) + '</strong>' +
    '<div class="class-students-grid"><div>' + teachersLabelHtml + teachersTable + '</div>' +
    '<div><div class="class-students-label">Учні</div>' + studentsTable + '</div>' +
    '<div><div class="class-students-label">Учні з журналу</div>' + journalTable + '</div></div></div></td>';
  row.after(tr);
  tr.querySelectorAll('.btn-find-student').forEach(btn => {
    if (btn.dataset.courseId) {
      btn.onclick = (e) => {
        e.stopPropagation();
        openFindStudentForClassroomModal(btn.dataset.name, btn.dataset.course, btn.dataset.courseId, row);
      };
    }
  });
  tr.querySelectorAll('.btn-add-teacher-to-course').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      openAddTeacherToCourseModal(btn.dataset.courseId, btn.dataset.courseName, btn.dataset.courseSection, row);
    };
  });
  tr.querySelectorAll('.btn-remove-teacher-from-course').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const email = btn.dataset.email;
      const cid = btn.dataset.courseId;
      if (!(await showConfirmModal(`Видалити викладача ${email} з курсу?`))) return;
      btn.disabled = true;
      try {
        const r = await fetch(`${API}/classroom/course/remove-teacher`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tokens, courseId: cid, email })
        });
        const text = await r.text();
        let data;
        try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
        if (data.error) throw new Error(data.error);
        if (courseRosterCache[cid]) {
          courseRosterCache[cid].teachers = (courseRosterCache[cid].teachers || []).filter(t => (t.email || '').toLowerCase() !== email.toLowerCase());
          setCourseRosterCache(cid, courseRosterCache[cid]);
        }
        const detailRow = row.nextElementSibling;
        if (detailRow?.classList?.contains('course-detail-row')) detailRow.remove();
        renderCourseRosterRow(row, row.dataset.name, row.dataset.section || '', courseRosterCache[cid]?.teachers || [], courseRosterCache[cid]?.students || []);
        renderDashboardMismatches();
        showAlertModal('Викладача видалено з курсу.');
      } catch (err) {
        showAlertModal('Помилка: ' + (err.message || err.name), 'Помилка', true);
      } finally {
        btn.disabled = false;
      }
    };
  });
  tr.querySelectorAll('.btn-remove-from-course').forEach(btn => {
    btn.onclick = async (e) => {
      e.stopPropagation();
      const email = btn.dataset.email;
      const cid = btn.dataset.courseId;
      if (!(await showConfirmModal(`Видалити ${email} з курсу?`))) return;
      btn.disabled = true;
      try {
        const r = await fetch(`${API}/classroom/course/remove-student`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tokens, courseId: cid, email })
        });
        const text = await r.text();
        let data;
        try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
        if (data.error) throw new Error(data.error);
        if (courseRosterCache[cid]) {
          courseRosterCache[cid].students = (courseRosterCache[cid].students || []).filter(st => st.email !== email);
          setCourseRosterCache(cid, courseRosterCache[cid]);
        }
        const detailRow = row.nextElementSibling;
        if (detailRow?.classList?.contains('course-detail-row')) detailRow.remove();
        renderCourseRosterRow(row, row.dataset.name, row.dataset.section || '', courseRosterCache[cid]?.teachers || [], courseRosterCache[cid]?.students || []);
        renderDashboardMismatches();
        showAlertModal('Учня видалено з курсу.');
      } catch (err) {
        showAlertModal('Помилка: ' + (err.message || err.name), 'Помилка', true);
      } finally {
        btn.disabled = false;
      }
    };
  });
}

async function toggleCourseRoster(row) {
  const courseId = row.dataset.id;
  const name = row.dataset.name;
  const detailRow = row.nextElementSibling;
  if (detailRow?.classList?.contains('course-detail-row')) {
    detailRow.remove();
    row.classList.remove('expanded');
    return;
  }
  if (row.classList.contains('loading')) return;
  let cached = courseRosterCache[courseId];
  if (!cached) cached = getCourseRosterCache(courseId);
  if (cached) {
    courseRosterCache[courseId] = cached;
    row.classList.add('expanded');
    renderCourseRosterRow(row, name, row.dataset.section || '', cached.teachers, cached.students);
    return;
  }
  if (courseRosterLoading.has(courseId)) return;
  courseRosterLoading.add(courseId);
  row.classList.add('loading', 'expanded');
  const origTd = row.querySelector('td');
  const origText = origTd?.textContent;
  if (origTd) origTd.textContent = 'Завантаження...';
  try {
    const r = await fetch(`${API}/classroom/course/roster`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, courseId })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    const teachers = data.teachers || [];
    const students = data.students || [];
    const roster = { teachers, students };
    courseRosterCache[courseId] = roster;
    setCourseRosterCache(courseId, roster);
    renderCourseRosterRow(row, name, row.dataset.section || '', teachers, students);
    renderDashboardMismatches();
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  } finally {
    courseRosterLoading.delete(courseId);
    row.classList.remove('loading');
    if (origTd) origTd.textContent = origText;
  }
}

async function loadClassroomCourses() {
  if (!tokens) return false;
  const status = document.getElementById('classroomCoursesStatus');
  status.style.display = 'block';
  status.className = 'status loading';
  status.textContent = 'Завантаження курсів...';
  try {
    const r = await fetch(`${API}/classroom/courses`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens })
    });
    const data = await r.json().catch(() => ({}));
    if (data.error) throw new Error(data.error);
    classroomCourses = data.courses || [];
    renderClassroomTable();
    saveToCache();
    status.className = 'status success';
    status.textContent = `Завантажено ${classroomCourses.length} курсів`;
    return true;
  } catch (e) {
    status.className = 'status error';
    status.textContent = 'Помилка: ' + (e.message || e.name);
    return false;
  }
}

function loadNzCreds() {
  const login = localStorage.getItem('nz_login');
  const password = localStorage.getItem('nz_password');
  const loginEl = document.getElementById('nzLogin');
  const passwordEl = document.getElementById('nzPassword');
  if (loginEl && login) loginEl.value = login;
  if (passwordEl && password) passwordEl.value = password;
  updateNzAuthUI();
}

document.getElementById('btnNzLogout').onclick = () => {
  localStorage.removeItem('nz_login');
  localStorage.removeItem('nz_password');
  document.getElementById('nzLogin').value = '';
  document.getElementById('nzPassword').value = '';
  clearCache();
  updateNzAuthUI();
};

document.getElementById('btnNzLogin').onclick = () => {
  const { login, password } = getNzCreds();
  if (!login || !password) {
    showAlertModal('Введіть логін та пароль', 'Увага');
    return;
  }
  saveNzCreds(login, password);
};

loadNzCreds();

document.querySelectorAll('.nav-item').forEach(el => {
  el.onclick = () => showPanel(el.dataset.panel);
});

async function loadJournals() {
  const { login, password } = getNzCreds();
  if (!login || !password) return false;
  const status = document.getElementById('parseStatus');
  status.style.display = 'block';
  status.className = 'status loading';
  status.textContent = 'Завантаження журналів з nz.ua...';
  try {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 120000);
    const r = await fetch(`${API}/parse`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ login, password }),
      signal: ctrl.signal
    });
    clearTimeout(t);
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Сервер повернув некоректні дані'); }
    if (data.error) throw new Error(data.error);
    parsedItems = data.items || [];
    saveNzCreds(login, password);
    status.className = 'status success';
    status.textContent = `Завантажено ${parsedItems.length} предметів`;
    renderJournalsTable();
    updateStats();
    saveToCache();
    return true;
  } catch (e) {
    status.className = 'status error';
    status.textContent = 'Помилка: ' + (e.message || e.name);
    return false;
  }
}

async function loadTeachers() {
  if (!tokens) return false;
  const ouName = 'Вчителі';
  try {
    const r = await fetch(`${API}/teachers/admin`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, ouName })
    });
    const data = await r.json().catch(() => ({}));
    if (data.error) return false;
    teachers = (data.teachers || []).slice().sort((a, b) => 
      (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
    renderTeachersTable();
    renderJournalTeachersTable();
    updateStats();
    saveToCache();
    return true;
  } catch (_) { return false; }
}

async function loadClasses() {
  if (!tokens) return false;
  const parentOuName = 'Учні';
  try {
    const r = await fetch(`${API}/classes/admin`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens: tokens, parentOuName })
    });
    const data = await r.json().catch(() => ({}));
    if (data.error) return false;
    adminClasses = data.classes || [];
    renderClassesTable();
    adminClassesCount = adminClasses.length;
    saveToCache();
    return true;
  } catch (_) { return false; }
}

async function loadAdminStats() {
  if (!tokens) return;
  try {
    const parentOuName = 'Учні';
    const r = await fetch(`${API}/classes/admin/stats`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, parentOuName })
    });
    const data = await r.json().catch(() => ({}));
    if (!data.error) {
      adminClassesCount = data.classesCount || 0;
      adminStudentsCount = data.studentsCount || 0;
      updateStats();
    }
  } catch (_) {}
}

async function autoLoadData(forceFetch = false) {
  const status = document.getElementById('parseStatus');
  const cache = loadFromCache();
  if (!forceFetch && cache.parsedItems?.length && cache.teachers?.length) {
    parsedItems = cache.parsedItems;
    teachers = cache.teachers;
    if (cache.adminStats) {
      adminClassesCount = cache.adminStats.adminClassesCount || 0;
      adminStudentsCount = cache.adminStats.adminStudentsCount || 0;
    }
    if (cache.adminClasses?.length) adminClasses = cache.adminClasses;
    status.style.display = 'none';
    document.getElementById('parseTable').style.display = 'block';
    renderJournalsTable();
    renderTeachersTable();
    renderJournalTeachersTable();
    if (adminClasses.length) renderClassesTable();
    updateStats();
    if (!adminClasses.length && loadTokens()) loadClasses().then(() => { saveToCache(); updateStats(); preloadClassStudents(); });
    if ((!adminClassesCount || !adminStudentsCount) && loadTokens()) loadAdminStats().then(() => saveToCache());
    preloadJournalDetails();
    if (adminClasses.length) preloadClassStudents();
    return;
  }
  status.style.display = 'block';
  status.className = 'status loading';
  status.textContent = 'Завантаження...';
  document.getElementById('parseTable').style.display = 'block';
  await Promise.all([loadJournals(), loadTeachers(), loadClasses(), loadAdminStats()]);
  saveToCache();
  updateStats();
  const j = parsedItems.length;
  const t = teachers.length;
  const c = adminClassesCount;
  const s = adminStudentsCount;
  status.className = 'status success';
  status.textContent = `Завантажено: ${j} предметів, ${t} вчителів, ${c} класів, ${s} учнів`;
  preloadJournalDetails();
  preloadClassStudents();
}

function preloadJournalDetails() {
  const journalRows = parsedItems.flatMap(({ subject, classes }) =>
    classes.map(c => ({ journalId: c.journalId, subgroupId: c.subgroupId }))
  );
  const toLoad = journalRows.filter(j => {
    const key = `${j.journalId}-${j.subgroupId || ''}`;
    return !getJournalDetailsCache(j.journalId, j.subgroupId) && !journalLoading.has(key);
  });
  if (!toLoad.length) return;
  const total = toLoad.length;
  toLoad.forEach(j => journalLoading.add(`${j.journalId}-${j.subgroupId || ''}`));
  const preloadEls = [document.getElementById('preloadStatus'), document.getElementById('preloadStatusDashboard')];
  let loaded = 0;
  const updatePreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.style.display = 'block';
      el.className = 'status loading';
      el.textContent = `Завантаження деталей журналів: ${loaded}/${total}`;
    });
  };
  const finishPreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.className = 'status success';
      el.textContent = `Завантажено всі журнали (${total})`;
    });
  };
  let idx = 0;
  const loadNext = async () => {
    while (idx < toLoad.length) {
      const j = toLoad[idx++];
      const { login, password } = getNzCreds();
      if (!login || !password) return;
      try {
        const r = await fetch(`${API}/parse/journal`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ journalId: j.journalId, subgroupId: j.subgroupId || undefined, login, password })
        });
        const data = await r.json().catch(() => ({}));
        if (!data.error) {
          setJournalDetailsCache(j.journalId, j.subgroupId, { teacher: data.teacher, students: data.students || [] });
          loaded++;
          updatePreload();
          updateStats();
          renderJournalTeachersTable();
        }
      } catch (_) {}
      journalLoading.delete(`${j.journalId}-${j.subgroupId || ''}`);
      await new Promise(r => setTimeout(r, 500));
    }
    finishPreload();
    updateStats();
    renderJournalTeachersTable();
  };
  updatePreload();
  loadNext();
}

function preloadClassStudents() {
  if (!tokens || !adminClasses.length) return;
  const toLoad = adminClasses.filter(c => {
    const path = c.orgUnitPath || '';
    return path && !getClassStudentsCache(path) && !classStudentsLoading.has(path);
  });
  if (!toLoad.length) return;
  const total = toLoad.length;
  toLoad.forEach(c => classStudentsLoading.add(c.orgUnitPath || ''));
  const preloadEls = [document.getElementById('preloadClassesStatus'), document.getElementById('preloadClassesStatusDashboard')];
  let loaded = 0;
  const updatePreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.style.display = 'block';
      el.className = 'status loading';
      el.textContent = `Завантаження учнів класів: ${loaded}/${total}`;
    });
  };
    const finishPreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.className = 'status success';
      el.textContent = `Завантажено учнів усіх класів (${total})`;
    });
    renderDashboardMismatches();
    document.querySelectorAll('.class-row.expanded').forEach(row => {
      const path = row.dataset.path;
      const name = row.dataset.name;
      const cached = getClassStudentsCache(path);
      const detailRow = row.nextElementSibling;
      if (detailRow?.classList?.contains('class-detail-row') && cached) {
        detailRow.remove();
        renderClassDetailRow(row, name, cached);
      }
    });
  };
  let idx = 0;
  const loadNext = async () => {
    while (idx < toLoad.length) {
      const c = toLoad[idx++];
      const path = c.orgUnitPath || '';
      try {
        const r = await fetch(`${API}/classes/admin/students`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tokens, orgUnitPath: path })
        });
        const data = await r.json().catch(() => ({}));
        if (!data.error) {
          const students = (data.students || []).slice().sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
          setClassStudentsCache(path, students);
          loaded++;
          updatePreload();
          renderDashboardMismatches();
          document.querySelectorAll('.class-row.expanded').forEach(row => {
            const p = row.dataset.path;
            const cached = getClassStudentsCache(p);
            const detailRow = row.nextElementSibling;
            if (detailRow?.classList?.contains('class-detail-row') && cached) {
              detailRow.remove();
              renderClassDetailRow(row, row.dataset.name, cached);
            }
          });
        }
      } catch (_) {}
      classStudentsLoading.delete(path);
      await new Promise(r => setTimeout(r, 300));
    }
    finishPreload();
  };
  updatePreload();
  loadNext();
}

function preloadCourseRosters() {
  if (!tokens || !classroomCourses.length) return;
  const toLoad = classroomCourses.filter(c => {
    const id = c.id || c.courseId;
    return id && !getCourseRosterCache(id) && !courseRosterLoading.has(id);
  });
  if (!toLoad.length) {
    renderDashboardMismatches();
    return;
  }
  const total = toLoad.length;
  toLoad.forEach(c => courseRosterLoading.add(c.id || c.courseId));
  const preloadEls = [document.getElementById('preloadClassroomStatus'), document.getElementById('preloadClassroomStatusDashboard')];
  let loaded = 0;
  const updatePreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.style.display = 'block';
      el.className = 'status loading';
      el.textContent = `Завантаження даних курсів: ${loaded}/${total}`;
    });
  };
  const finishPreload = () => {
    preloadEls.forEach(el => {
      if (!el) return;
      el.className = 'status success';
      el.textContent = `Завантажено дані всіх курсів (${total})`;
    });
    renderDashboardMismatches();
  };
  let idx = 0;
  const loadNext = async () => {
    while (idx < toLoad.length) {
      const c = toLoad[idx++];
      const courseId = c.id || c.courseId;
      try {
        const r = await fetch(`${API}/classroom/course/roster`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tokens, courseId })
        });
        const data = await r.json().catch(() => ({}));
        if (!data.error) {
          const teachers = data.teachers || [];
          const students = data.students || [];
          const roster = { teachers, students };
          courseRosterCache[courseId] = roster;
          setCourseRosterCache(courseId, roster);
          loaded++;
          updatePreload();
          renderDashboardMismatches();
        }
      } catch (_) {}
      courseRosterLoading.delete(courseId);
      await new Promise(r => setTimeout(r, 300));
    }
    finishPreload();
  };
  updatePreload();
  loadNext();
}

(async function initAutoLoad() {
  await new Promise(r => setTimeout(r, 300));
  if (!loadTokens()) return;
  const cache = loadFromCache();
  if (cache.classroomCourses?.length) {
    if (!classroomCourses.length) classroomCourses = cache.classroomCourses;
    classroomCourses.forEach(c => {
      const id = c.id || c.courseId;
      const roster = getCourseRosterCache(id);
      if (roster) courseRosterCache[id] = roster;
    });
    renderClassroomTable();
    preloadCourseRosters();
  } else {
    loadClassroomCourses().then(ok => { if (ok) preloadCourseRosters(); });
  }
  if (!localStorage.getItem('nz_login') || !localStorage.getItem('nz_password')) return;
  await autoLoadData(false);
})();

function clearJournalDetailsCache() {
  localStorage.removeItem(CACHE_KEYS.journalDetails);
  journalLoading.clear();
}

async function handleParse() {
  showPanel('journals');
  const { login, password } = getNzCreds();
  if (!login || !password) {
    showAlertModal('Увійдіть в nz.ua або введіть логін та пароль', 'Увага');
    return;
  }
  const confirmed = await showConfirmModal(
    'Повне оновлення журналів включає завантаження списку предметів та деталей кожного журналу (викладач, учні). Це може зайняти кілька хвилин залежно від кількості журналів. Продовжити?',
    'Оновлення журналів'
  );
  if (!confirmed) return;
  const btn = document.getElementById('btnParse');
  const status = document.getElementById('parseStatus');
  const progressRow = document.getElementById('parseProgressRow');
  const progressText = document.getElementById('parseProgressText');
  const progressBar = document.getElementById('parseProgressBar');
  const preloadEl = document.getElementById('preloadStatus');
  btn.disabled = true;
  status.style.display = 'block';
  progressRow.style.display = 'block';
  if (preloadEl) preloadEl.style.display = 'none';
  clearJournalDetailsCache();
  try {
    progressText.textContent = 'Крок 1/2: Завантаження списку журналів...';
    progressBar.style.width = '5%';
    status.className = 'status loading';
    status.textContent = 'Вхід на nz.ua, отримання списку предметів (15–30 сек)...';
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), 120000);
    const r = await fetch(`${API}/parse`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ login, password }),
      signal: ctrl.signal
    });
    clearTimeout(t);
    const text = await r.text();
    if (!text) throw new Error('Порожня відповідь. Перевірте: 1) npm run dev запущений 2) backend на порту 3003');
    let data;
    try {
      data = JSON.parse(text);
    } catch (_) {
      throw new Error('Сервер повернув некоректні дані. Запустіть: npm run dev');
    }
    if (data.error) throw new Error(data.error);
    parsedItems = data.items || [];
    saveNzCreds(login, password);
    saveToCache();
    renderJournalsTable();
    const journalRows = parsedItems.flatMap(({ subject, classes }) =>
      classes.map(c => ({ subject, journalId: c.journalId, subgroupId: c.subgroupId }))
    );
    const total = journalRows.length;
    if (total === 0) {
      progressBar.style.width = '100%';
      progressText.textContent = 'Готово';
      status.className = 'status success';
      status.textContent = `Завантажено ${parsedItems.length} предметів (журналів: 0)`;
      updateStats();
      return;
    }
    progressBar.style.width = '10%';
    progressText.textContent = `Крок 2/2: Деталі журналів 0 / ${total}`;
    status.textContent = 'Завантаження деталей (викладач, учні) для кожного журналу...';
    let loaded = 0;
    const updateProgress = () => {
      const pct = 10 + Math.round(90 * loaded / total);
      progressBar.style.width = pct + '%';
      progressText.textContent = `Крок 2/2: Деталі журналів ${loaded} / ${total}`;
      status.textContent = `Завантажено деталей: ${loaded} з ${total} журналів`;
    };
    for (const j of journalRows) {
      try {
        const resp = await fetch(`${API}/parse/journal`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ journalId: j.journalId, subgroupId: j.subgroupId || undefined, login, password })
        });
        const jData = await resp.json().catch(() => ({}));
        if (!jData.error) {
          setJournalDetailsCache(j.journalId, j.subgroupId, { teacher: jData.teacher, students: jData.students || [] });
        }
      } catch (_) {}
      loaded++;
      updateProgress();
      updateStats();
      renderJournalTeachersTable();
      await new Promise(r => setTimeout(r, 300));
    }
    progressBar.style.width = '100%';
    progressText.textContent = `Готово: ${total} журналів`;
    status.className = 'status success';
    status.textContent = `Повне оновлення завершено: ${parsedItems.length} предметів, ${total} журналів з деталями`;
    updateStats();
    renderJournalTeachersTable();
  } catch (e) {
    status.className = 'status error';
    const msg = e.name === 'AbortError' ? 'Час очікування вийшов (2 хв). Спробуйте ще раз.' : e.message;
    status.textContent = 'Помилка: ' + msg;
    progressText.textContent = 'Помилка';
  } finally {
    btn.disabled = false;
    progressRow.style.display = 'none';
  }
}
document.getElementById('btnParse').onclick = handleParse;

function escapeHtml(s) {
  const d = document.createElement('div');
  d.textContent = s || '';
  return d.innerHTML;
}

function showAlertModal(message, title = 'Повідомлення', isError = false) {
  const modal = document.getElementById('modalAlert');
  const titleEl = document.getElementById('modalAlertTitle');
  const messageEl = document.getElementById('modalAlertMessage');
  const modalInner = modal?.querySelector('.modal');
  if (!modal || !titleEl || !messageEl) return;
  titleEl.textContent = title || 'Повідомлення';
  messageEl.textContent = message || '';
  if (modalInner) modalInner.classList.toggle('modal-error', !!isError);
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) { modal.style.display = 'none'; modal.onclick = null; } };
  document.getElementById('btnModalAlertOk').onclick = () => { modal.style.display = 'none'; modal.onclick = null; };
}

function showConfirmModal(message, title = 'Підтвердження', isError = false) {
  return new Promise(resolve => {
    const modal = document.getElementById('modalConfirm');
    const modalInner = modal?.querySelector('.modal');
    const titleEl = document.getElementById('modalConfirmTitle');
    const messageEl = document.getElementById('modalConfirmMessage');
    if (!modal || !titleEl || !messageEl) { resolve(false); return; }
    titleEl.textContent = title || 'Підтвердження';
    messageEl.textContent = message || '';
    if (modalInner) modalInner.classList.toggle('modal-error', !!isError);
    modal.style.display = 'flex';
    const close = (result) => {
      modal.style.display = 'none';
      modal.onclick = null;
      document.getElementById('btnModalConfirmOk').onclick = null;
      document.getElementById('btnModalConfirmCancel').onclick = null;
      resolve(result);
    };
    modal.onclick = (e) => { if (e.target === modal) close(false); };
    document.getElementById('btnModalConfirmOk').onclick = () => close(true);
    document.getElementById('btnModalConfirmCancel').onclick = () => close(false);
  });
}

function sortJournalRows(rows) {
  const parseClass = (s) => {
    const norm = normalizeClassLabel(s || '');
    const m = norm.match(/^(\d+)[-\s]*(.*)$/);
    return m ? { grade: parseInt(m[1], 10), letter: (m[2] || '').trim() } : { grade: 0, letter: norm };
  };
  return rows.slice().sort((a, b) => {
    const subj = (a.subject || '').localeCompare(b.subject || '', 'uk');
    if (subj !== 0) return subj;
    const pa = parseClass(a.classLabel);
    const pb = parseClass(b.classLabel);
    if (pa.grade !== pb.grade) return pb.grade - pa.grade;
    return (pa.letter || '').localeCompare(pb.letter || '', 'uk');
  });
}

function renderJournalDetailRow(row, teacher, students) {
  const sorted = (students || []).slice().sort((a, b) => (a.name || '').localeCompare(b.name || '', 'uk'));
  const tr = document.createElement('tr');
  tr.className = 'journal-detail-row';
  let teacherHtml = teacher ? `<p><strong>Викладач:</strong> ${escapeHtml(teacher.name)}</p>` : '<p><strong>Викладач:</strong> —</p>';
  let studentsHtml = sorted.length ? '<table><thead><tr><th>#</th><th>ПІБ учня</th></tr></thead><tbody>' +
    sorted.map((s, i) => `<tr><td>${i + 1}</td><td>${escapeHtml(s.name)}</td></tr>`).join('') +
    '</tbody></table>' : '<p style="color:var(--text-muted)">Немає учнів</p>';
  tr.innerHTML = '<td colspan="2"><div class="journal-details">' + teacherHtml + '<div class="journal-students"><strong>Учні:</strong>' + studentsHtml + '</div></div></td>';
  row.after(tr);
}

async function toggleJournalDetails(row) {
  const journalId = row.dataset.journalId;
  const subgroupId = row.dataset.subgroupId || '';
  const detailRow = row.nextElementSibling;
  if (detailRow?.classList?.contains('journal-detail-row')) {
    detailRow.remove();
    row.classList.remove('expanded');
    return;
  }
  if (row.classList.contains('loading')) return;
  const cached = getJournalDetailsCache(journalId, subgroupId);
  if (cached) {
    row.classList.add('expanded');
    renderJournalDetailRow(row, cached.teacher, cached.students);
    return;
  }
  const loadKey = `${journalId}-${subgroupId || ''}`;
  if (journalLoading.has(loadKey)) {
    row.classList.add('loading', 'expanded');
    const origTd = row.querySelector('td:last-child');
    const origText = origTd?.textContent;
    if (origTd) origTd.textContent = 'Завантаження...';
    let attempts = 0;
    const checkInterval = setInterval(() => {
      attempts++;
      if (!journalLoading.has(loadKey) || attempts > 200) {
        clearInterval(checkInterval);
        const c = getJournalDetailsCache(journalId, subgroupId);
        row.classList.remove('loading');
        if (origTd) origTd.textContent = origText;
        if (c) renderJournalDetailRow(row, c.teacher, c.students);
      }
    }, 300);
    return;
  }
  journalLoading.add(loadKey);
  row.classList.add('loading', 'expanded');
  const origTd = row.querySelector('td:last-child');
  const origText = origTd?.textContent;
  if (origTd) origTd.textContent = 'Завантаження...';
  try {
    const { login, password } = getNzCreds();
    if (!login || !password) throw new Error('Увійдіть в nz.ua');
    const r = await fetch(`${API}/parse/journal`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ journalId, subgroupId: subgroupId || undefined, login, password })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    const teacher = data.teacher;
    const students = data.students || [];
    setJournalDetailsCache(journalId, subgroupId, { teacher, students });
    renderJournalDetailRow(row, teacher, students);
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  } finally {
    journalLoading.delete(loadKey);
    row.classList.remove('loading');
    if (origTd) origTd.textContent = origText;
  }
}

async function moveStudentToClass(email, path, fromPath, classRow) {
  if (!loadTokens() || !email || !path) return;
  if (!(await showConfirmModal(`Перемістити учня ${email} в цей клас?`))) return;
  try {
    const r = await fetch(`${API}/users/move`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, email, orgUnitPath: path })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    const fromCache = fromPath ? getClassStudentsCache(fromPath) : null;
    const moved = fromCache?.find(s => s.email === email);
    const targetCache = getClassStudentsCache(path);
    if (targetCache && moved) {
      targetCache.push(moved);
      targetCache.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
      setClassStudentsCache(path, targetCache);
    }
    if (fromCache) {
      const filtered = fromCache.filter(s => s.email !== email);
      setClassStudentsCache(fromPath, filtered);
    }
    const detailRow = classRow.nextElementSibling;
    if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
    renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
    const otherRow = document.querySelector(`.class-row[data-path="${fromPath}"]`);
    if (otherRow?.nextElementSibling?.classList?.contains('class-detail-row')) {
      otherRow.nextElementSibling.remove();
      renderClassDetailRow(otherRow, otherRow.dataset.name, getClassStudentsCache(fromPath));
    }
    renderDashboardMismatches();
    showAlertModal('Учня переміщено.');
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
}

async function moveStudentToGraduates(email, path, classRow) {
  if (!loadTokens() || !email) return;
  if (!(await showConfirmModal(`Перемістити учня ${email} в підрозділ «Випускники»?`))) return;
  try {
    const r = await fetch(`${API}/users/move-to-graduates`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, email })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    const cached = getClassStudentsCache(path);
    if (cached) {
      const filtered = cached.filter(s => s.email !== email);
      setClassStudentsCache(path, filtered);
    }
    const detailRow = classRow.nextElementSibling;
    if (detailRow?.classList?.contains('class-detail-row')) detailRow.remove();
    renderClassDetailRow(classRow, classRow.dataset.name, getClassStudentsCache(path));
    renderDashboardMismatches();
    showAlertModal('Учня переміщено в Випускники.');
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
}

function parseStudentName(fullName) {
  const parts = (fullName || '').trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) return { familyName: parts[0], givenName: parts.slice(1).join(' ') };
  if (parts.length === 1) return { familyName: parts[0], givenName: '' };
  return { familyName: '', givenName: '' };
}

function formatTeacherDisplayName(name) {
  if (!name || typeof name !== 'string') return '';
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length >= 2) return [parts[parts.length - 1], parts.slice(0, -1).join(' ')].filter(Boolean).join(' ');
  return name.trim();
}

function renderClassDetailRow(row, name, students) {
  const journalStudents = getStudentsFromJournalsForClass(name);
  const { googleKeys, journalKeys, journalKeysCorrect } = getStudentMatchSets(students, journalStudents);
  const orgUnitPath = row.dataset.path || '';
  const tr = document.createElement('tr');
  tr.className = 'class-detail-row';
  const toSwapStudents = [];
  const googleRows = (students || []).map((s, i) => {
    const fn = (s.familyName || '').trim();
    const gn = (s.givenName || '').trim();
    const key1 = normalizeNameForMatch(`${fn} ${gn}`);
    const key2 = normalizeNameForMatch(`${gn} ${fn}`);
    const missing = !journalKeys.has(key1) && !journalKeys.has(key2) && journalKeys.size > 0;
    const needsSwap = !missing && (journalKeys.has(key1) || journalKeys.has(key2)) && !journalKeysCorrect.has(key1) && journalKeysCorrect.size > 0;
    if (needsSwap) toSwapStudents.push({ email: s.email, familyName: s.familyName, givenName: s.givenName });
    const cls = missing ? ' class="teacher-missing"' : (needsSwap ? ' class="name-swapped"' : '');
    const moveBtn = missing ? `<button type="button" class="btn-move-to-graduates btn-secondary" data-email="${escapeHtml(s.email)}" data-path="${escapeHtml(orgUnitPath)}" title="Перемістити в Випускники">×</button>` : '';
    const editBtn = `<button type="button" class="btn-edit-user btn-secondary" data-email="${escapeHtml(s.email)}" data-family="${escapeHtml(s.familyName || '')}" data-given="${escapeHtml(s.givenName || '')}" data-type="student" data-path="${escapeHtml(orgUnitPath)}" title="Редагувати">✏️</button>`;
    const swapBtn = needsSwap ? `<button type="button" class="btn-swap-names btn-secondary" data-email="${escapeHtml(s.email)}" data-family="${escapeHtml(s.familyName || '')}" data-given="${escapeHtml(s.givenName || '')}" data-type="student" data-path="${escapeHtml(orgUnitPath)}" title="Поміняти місцями ім'я та прізвище">⇄</button>` : '';
    return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(s.familyName)}</td><td>${escapeHtml(s.givenName)}</td><td>${escapeHtml(s.email)}</td><td class="btn-cell"><div class="btn-cell-inner">${editBtn}${swapBtn}${moveBtn}</div></td></tr>`;
  });
  const googleHeaderTh = toSwapStudents.length > 3 ? `<th><button type="button" class="btn-mass-swap-students btn-secondary btn-sm" title="Поміняти місцями ім'я та прізвище у всіх">⇄ Усіх (${toSwapStudents.length})</button></th>` : '<th></th>';
  const googleTable = (students || []).length ? '<table class="class-students-table google-table"><thead><tr><th>#</th><th>Прізвище</th><th>Ім\'я</th><th>Email</th>' + googleHeaderTh + '</tr></thead><tbody>' + googleRows.join('') + '</tbody></table>' : '<p style="color:var(--text-muted);margin:0;">Немає учнів</p>';
  const journalTable = journalStudents.length ? '<table class="class-students-table journal-table"><thead><tr><th>#</th><th>ПІБ учня</th><th></th></tr></thead><tbody>' +
    journalStudents.map((studentName, i) => {
      const key = normalizeNameForMatch(studentName);
      const missing = !googleKeys.has(key) && googleKeys.size > 0;
      const inOther = missing ? getStudentInOtherClass(studentName, orgUnitPath) : null;
      const cls = inOther ? ' class="student-in-other-class"' : (missing ? ' class="teacher-missing"' : '');
      let btn = '';
      if (inOther) {
        btn = `<button type="button" class="btn-move-to-class btn-secondary" data-email="${escapeHtml(inOther.email)}" data-path="${escapeHtml(orgUnitPath)}" data-from-path="${escapeHtml(inOther.fromPath || '')}" title="Перемістити в цей клас">↗</button>`;
      } else if (missing) {
        btn = `<button type="button" class="btn-find-student btn-secondary" data-name="${escapeHtml(studentName)}" data-class="${escapeHtml(name)}" data-path="${escapeHtml(orgUnitPath)}" title="Знайти в Google Admin">🔍</button>`;
      }
      return `<tr${cls}><td>${i + 1}</td><td>${escapeHtml(studentName)}</td><td class="btn-cell"><div class="btn-cell-inner">${btn}</div></td></tr>`;
    }).join('') + '</tbody></table>' : (areAllJournalDetailsLoaded() ? '<p style="color:var(--text-muted);margin:0;">Немає даних</p>' : '<p class="status loading" style="margin:0;">Завантаження...</p>');
  tr.innerHTML = '<td colspan="2"><div class="class-students"><strong>Учні ' + escapeHtml(name) + ':</strong>' +
    '<div class="class-students-grid">' +
    '<div><div class="class-students-label">Google Admin</div>' + googleTable + '</div>' +
    '<div><div class="class-students-label">Журнали</div>' + journalTable + '</div>' +
    '</div></div></td>';
  row.after(tr);
  tr.querySelectorAll('.btn-find-student').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      openFindStudentModal(btn.dataset.name, btn.dataset.class, btn.dataset.path);
    };
  });
  tr.querySelectorAll('.btn-move-to-class').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      moveStudentToClass(btn.dataset.email, btn.dataset.path, btn.dataset.fromPath, row);
    };
  });
  tr.querySelectorAll('.btn-edit-user').forEach(btn => {
    if (btn.dataset.type === 'student') {
      btn.onclick = (e) => {
        e.stopPropagation();
        openEditUserModal(btn.dataset.email, btn.dataset.family, btn.dataset.given, 'student', btn.dataset.path);
      };
    }
  });
  tr.querySelectorAll('.btn-swap-names').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      swapUserNames(btn.dataset.email, btn.dataset.family, btn.dataset.given, 'student', btn.dataset.path);
    };
  });
  tr.querySelectorAll('.btn-move-to-graduates').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      moveStudentToGraduates(btn.dataset.email, btn.dataset.path, row);
    };
  });
  tr.querySelectorAll('.btn-mass-swap-students').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      swapUserNamesBulk(toSwapStudents, 'student', orgUnitPath);
    };
  });
}

async function toggleClassStudents(row) {
  const path = row.dataset.path;
  const name = row.dataset.name;
  const detailRow = row.nextElementSibling;
  if (detailRow?.classList?.contains('class-detail-row')) {
    detailRow.remove();
    row.classList.remove('expanded');
    return;
  }
  if (row.classList.contains('loading')) return;
  const cached = getClassStudentsCache(path);
  if (cached) {
    row.classList.add('expanded');
    renderClassDetailRow(row, name, cached);
    preloadClassStudents();
    return;
  }
  if (classStudentsLoading.has(path)) {
    row.classList.add('loading', 'expanded');
    const origTd = row.querySelector('td');
    const origText = origTd?.textContent;
    if (origTd) origTd.textContent = 'Завантаження...';
    let attempts = 0;
    const checkInterval = setInterval(() => {
      attempts++;
      if (!classStudentsLoading.has(path) || attempts > 200) {
        clearInterval(checkInterval);
        const c = getClassStudentsCache(path);
        row.classList.remove('loading');
        if (origTd) origTd.textContent = origText;
        if (c) renderClassDetailRow(row, name, c);
      }
    }, 300);
    return;
  }
  classStudentsLoading.add(path);
  row.classList.add('loading', 'expanded');
  const origTd = row.querySelector('td');
  const origText = origTd?.textContent;
  if (origTd) origTd.textContent = 'Завантаження...';
  try {
    const r = await fetch(`${API}/classes/admin/students`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, orgUnitPath: path })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    const students = (data.students || []).slice().sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
    setClassStudentsCache(path, students);
    renderClassDetailRow(row, name, students);
    renderDashboardMismatches();
    preloadClassStudents();
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  } finally {
    classStudentsLoading.delete(path);
    row.classList.remove('loading');
    if (origTd) origTd.textContent = origText;
  }
}

function openAddTeacherModal() {
  const modal = document.getElementById('modalAddTeacher');
  const fn = document.getElementById('addTeacherFamilyName');
  const gn = document.getElementById('addTeacherGivenName');
  const email = document.getElementById('addTeacherEmail');
  const pwd = document.getElementById('addTeacherPassword');
  fn.value = '';
  gn.value = '';
  email.value = '';
  pwd.value = generateTeacherPassword();
  fn.oninput = gn.oninput = () => {
    const f = fn.value.trim();
    const g = gn.value.trim();
    email.value = f || g ? generateTeacherEmail(f, g) : '';
  };
  modal.style.display = 'flex';
  modal.onclick = (e) => { if (e.target === modal) closeAddTeacherModal(); };
}

function closeAddTeacherModal() {
  document.getElementById('modalAddTeacher').style.display = 'none';
}

document.getElementById('btnAddTeacher').onclick = () => {
  if (!loadTokens()) { showAlertModal('Спочатку увійдіть через Google', 'Увага'); return; }
  openAddTeacherModal();
};

document.getElementById('btnModalAddTeacherCancel').onclick = closeAddTeacherModal;

document.getElementById('btnModalAddTeacherSubmit').onclick = async () => {
  const fn = document.getElementById('addTeacherFamilyName').value.trim();
  const gn = document.getElementById('addTeacherGivenName').value.trim();
  const email = document.getElementById('addTeacherEmail').value.trim();
  const pwd = document.getElementById('addTeacherPassword').value;
  if (!fn || !gn) { showAlertModal('Введіть прізвище та ім\'я', 'Увага'); return; }
  if (!email) { showAlertModal('Email не згенеровано', 'Увага'); return; }
  try {
    const r = await fetch(`${API}/teachers/create`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, givenName: gn, familyName: fn, primaryEmail: email, password: pwd })
    });
    const text = await r.text();
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Помилка відповіді'); }
    if (data.error) throw new Error(data.error);
    closeAddTeacherModal();
    teachers.push({ familyName: fn, givenName: gn, email });
    teachers.sort((a, b) => (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
    renderTeachersTable();
    renderJournalTeachersTable();
    updateStats();
    saveToCache();
    showAlertModal(`Викладача створено.\nEmail: ${email}\nПароль: ${pwd}\n\nЗбережіть пароль!`);
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  }
};

document.getElementById('btnLoadAdminTeachers').onclick = async () => {
  showPanel('teachers');
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const ouName = 'Вчителі';
  const btn = document.getElementById('btnLoadAdminTeachers');
  btn.disabled = true;
  btn.textContent = 'Завантаження...';
  try {
    const r = await fetch(`${API}/teachers/admin`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, ouName })
    });
    const text = await r.text();
    if (!text) throw new Error('Порожня відповідь сервера');
    let data;
    try { data = JSON.parse(text); } catch (_) { throw new Error('Сервер повернув некоректні дані'); }
    if (data.error) throw new Error(data.error);
    teachers = (data.teachers || []).slice().sort((a, b) => 
      (a.familyName || '').localeCompare(b.familyName || '', 'uk'));
    renderTeachersTable();
    renderJournalTeachersTable();
    updateStats();
    saveToCache();
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Оновити';
  }
};

document.getElementById('btnLoadAdminClasses').onclick = async () => {
  showPanel('classes');
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const parentOuName = 'Учні';
  const btn = document.getElementById('btnLoadAdminClasses');
  btn.disabled = true;
  btn.textContent = 'Завантаження...';
  try {
    const r = await fetch(`${API}/classes/admin`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens: tokens, parentOuName })
    });
    const text = await r.text();
    if (!text) throw new Error('Порожня відповідь сервера');
    let data;
    try { data = JSON.parse(text); } catch (_) {
      if (text.includes('Cannot POST') || text.includes('404')) throw new Error('Маршрут не знайдено. Перезапустіть: npm run dev');
      throw new Error('Сервер повернув некоректні дані');
    }
    if (data.error) throw new Error(data.error);
    adminClasses = data.classes || [];
    renderClassesTable();
    adminClassesCount = adminClasses.length;
    const r2 = await fetch(`${API}/classes/admin/stats`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, parentOuName })
    });
    const d2 = await r2.json().catch(() => ({}));
    if (!d2.error) adminStudentsCount = d2.studentsCount || 0;
    updateStats();
    saveToCache();
    preloadClassStudents();
  } catch (e) {
    showAlertModal('Помилка: ' + e.message, 'Помилка', true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Оновити';
  }
};

document.getElementById('btnLoadClassroomCourses').onclick = async () => {
  showPanel('classroom');
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const ok = await loadClassroomCourses();
  if (ok) preloadCourseRosters();
};

document.getElementById('btnAddTeacherToSelectedCourses').onclick = async () => {
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const listEl = document.getElementById('classroomCoursesList');
  const checkboxes = listEl ? [...listEl.querySelectorAll('.course-checkbox:checked')] : [];
  const courseIds = checkboxes.map(cb => cb.dataset.courseId).filter(Boolean);
  const courseRows = checkboxes.map(cb => cb.closest('.course-row')).filter(Boolean);
  if (!courseIds.length) {
    showAlertModal('Оберіть курси для додавання викладача', 'Увага');
    return;
  }
  openAddTeacherToMassCoursesModal(courseIds, courseRows);
};

document.getElementById('btnArchiveSelectedCourses').onclick = async () => {
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const listEl = document.getElementById('classroomCoursesList');
  const checked = listEl ? [...listEl.querySelectorAll('.course-checkbox:checked')].map(cb => cb.dataset.courseId).filter(Boolean) : [];
  if (!checked.length) {
    showAlertModal('Оберіть курси для архівування', 'Увага');
    return;
  }
  if (!(await showConfirmModal(`Архівувати ${checked.length} обраних курсів?`))) return;
  const btn = document.getElementById('btnArchiveSelectedCourses');
  btn.disabled = true;
  btn.textContent = 'Архівування...';
  try {
    const r = await fetch(`${API}/classroom/courses/archive`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, courseIds: checked })
    });
    const data = await r.json().catch(() => ({}));
    if (data.error) throw new Error(data.error);
    const results = data.results || [];
    const okCount = results.filter(x => x.success).length;
    const errCount = results.filter(x => !x.success).length;
    checked.forEach(cid => {
      delete courseRosterCache[cid];
      try {
        const all = localStorage.getItem(CACHE_KEYS.courseRosters) || '{}';
        const obj = JSON.parse(all);
        delete obj[cid];
        localStorage.setItem(CACHE_KEYS.courseRosters, JSON.stringify(obj));
      } catch (_) {}
    });
    const ok = await loadClassroomCourses();
    if (ok) preloadCourseRosters();
    renderDashboardMismatches();
    if (errCount) {
      showAlertModal(`Архівовано ${okCount} курсів. Помилок: ${errCount}`, 'Результат', true);
    } else {
      showAlertModal(`Архівовано ${okCount} курсів.`);
    }
  } catch (e) {
    showAlertModal('Помилка: ' + (e.message || e.name), 'Помилка', true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Архівувати обрані курси';
  }
};

let syncCancelled = false;

document.getElementById('btnSync').onclick = async () => {
  showPanel('manage');
  if (!loadTokens()) {
    showAlertModal('Спочатку увійдіть через Google', 'Увага');
    return;
  }
  const { adminMismatches } = getMismatchClasses();
  if (adminMismatches.length > 0) {
    showAlertModal('У Класах з Google Admin та Вчителі не має бути невідповідностей. Виправте невідповідності на Головній перед створенням курсів.', 'Увага', true);
    return;
  }
  const { login, password } = getNzCreds();
  if (!login || !password) {
    showAlertModal('Введіть логін та пароль nz.ua', 'Увага');
    return;
  }
  const ok = await showConfirmModal('Запустити створення курсів у Google Classroom та заповнення учнями та вчителями?', 'Підтвердження', true);
  if (!ok) return;
  syncCancelled = false;
  const btn = document.getElementById('btnSync');
  const status = document.getElementById('syncStatus');
  const logEl = document.getElementById('syncLog');
  const progressRow = document.getElementById('syncProgressRow');
  const progressText = document.getElementById('syncProgressText');
  const progressBar = document.getElementById('syncProgressBar');
  const stopBtn = document.getElementById('btnSyncStop');
  btn.disabled = true;
  status.style.display = 'block';
  status.className = 'status loading';
  status.textContent = 'Підготовка...';
  progressRow.style.display = 'block';
  progressText.textContent = '0 / 0';
  progressBar.style.width = '0%';
  logEl.style.display = 'block';
  logEl.innerHTML = '';
  try {
    const initRes = await fetch(`${API}/sync/init`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tokens, teacherMapping: [], login, password, ouName: 'Вчителі' })
    });
    const initData = await initRes.json().catch(() => ({}));
    if (initData.error) throw new Error(initData.error);
    const journals = initData.journals || [];
    if (!journals.length) {
      status.className = 'status';
      status.textContent = 'Журналів для створення не знайдено';
      progressRow.style.display = 'none';
      return;
    }
    const total = journals.length;
    let done = 0;
    const updateProgress = () => {
      progressText.textContent = `${done} / ${total}`;
      progressBar.style.width = total ? (100 * done / total) + '%' : '0%';
    };
    stopBtn.onclick = () => { syncCancelled = true; };
    for (const j of journals) {
      if (syncCancelled) break;
      status.textContent = `Створення: ${j.subject} — ${j.classLabel}...`;
      try {
        const r = await fetch(`${API}/sync/create-one`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            tokens,
            login,
            password,
            journal: j,
            emailMapArr: initData.emailMap || [],
            adminClasses: initData.adminClasses || [],
            currentUserEmail: initData.currentUserEmail
          })
        });
        const x = await r.json().catch(() => ({}));
        if (x.error) {
          const div = document.createElement('div');
          div.className = 'log-entry err';
          div.textContent = `${x.courseName || j.subject + ' - ' + j.classLabel}: ${x.error}`;
          logEl.appendChild(div);
        } else {
          const div = document.createElement('div');
          div.className = 'log-entry ok';
          const extra = x.studentsAdded != null ? ` (учнів: ${x.studentsAdded})` : '';
          div.textContent = `${x.courseName} ✓${extra}`;
          logEl.appendChild(div);
        }
      } catch (_) {}
      done++;
      updateProgress();
      logEl.scrollTop = logEl.scrollHeight;
    }
    status.className = syncCancelled ? 'status' : 'status success';
    status.textContent = syncCancelled ? `Зупинено. Створено ${done} з ${total}` : `Готово! Створено ${done} курсів`;
    if (done) {
      classroomCourses = [];
      loadClassroomCourses().then(ok => { if (ok) preloadCourseRosters(); });
      renderClassroomTable();
      renderDashboardMismatches();
    }
  } catch (e) {
    status.className = 'status error';
    status.textContent = 'Помилка: ' + e.message;
    logEl.innerHTML = `<div class="log-entry err">${escapeHtml(e.message)}</div>`;
  } finally {
    btn.disabled = false;
    progressRow.style.display = 'none';
    stopBtn.onclick = null;
  }
};
