import { google } from 'googleapis';

function getOAuth2Client(tokens) {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI || 'http://localhost:5173/callback'
  );
  if (tokens) oauth2Client.setCredentials(tokens);
  return oauth2Client;
}

export function getClassroomClient(tokens) {
  const auth = getOAuth2Client(tokens);
  return google.classroom({ version: 'v1', auth });
}

export async function getTeachersFromAdmin(tokens, ouName = 'Вчителі') {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  let orgUnitPath = `/${ouName}`;
  try {
    const orgRes = await admin.orgunits.list({
      customerId: 'my_customer'
    });
    const ou = orgRes.data.organizationUnits?.find(
      o => o.name === ouName || o.orgUnitPath?.endsWith(`/${ouName}`)
    );
    if (ou) orgUnitPath = ou.orgUnitPath;
  } catch (_) {}
  const teachers = [];
  let pageToken;
  do {
    const res = await admin.users.list({
      customer: 'my_customer',
      query: `orgUnitPath='${orgUnitPath}'`,
      maxResults: 500,
      pageToken,
      viewType: 'admin_view',
      orderBy: 'familyName',
      sortOrder: 'ASCENDING'
    });
    for (const u of res.data.users || []) {
      const name = [u.name?.givenName, u.name?.familyName].filter(Boolean).join(' ');
      if (u.primaryEmail && name) {
        teachers.push({
          email: u.primaryEmail,
          name: name.trim(),
          givenName: u.name?.givenName,
          familyName: u.name?.familyName
        });
      }
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return teachers;
}

async function getOrgUnitChildrenRecursive(admin, customerId, parentPath, classes, maxDepth = 5) {
  if (maxDepth <= 0) return;
  try {
    const res = await admin.orgunits.list({
      customerId,
      orgUnitPath: parentPath,
      type: 'CHILDREN'
    });
    for (const o of res.data.organizationUnits || []) {
      classes.push({ name: o.name, orgUnitPath: o.orgUnitPath, orgUnitId: o.orgUnitId });
      await getOrgUnitChildrenRecursive(admin, customerId, o.orgUnitPath, classes, maxDepth - 1);
    }
  } catch (_) {}
}

export async function getClassesFromAdmin(tokens, parentOuName = 'Учні') {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  let parentPath = `/${parentOuName}`;
  try {
    const orgRes = await admin.orgunits.list({
      customerId: 'my_customer'
    });
    const ou = orgRes.data.organizationUnits?.find(
      o => o.name === parentOuName || o.orgUnitPath?.endsWith(`/${parentOuName}`)
    );
    if (ou) parentPath = ou.orgUnitPath;
  } catch (_) {}
  const allClasses = [];
  await getOrgUnitChildrenRecursive(admin, 'my_customer', parentPath, allClasses);
  return allClasses;
}

export async function createUserInOrgUnit(tokens, { givenName, familyName, primaryEmail, password }, ouName = 'Вчителі') {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  let orgUnitPath = `/${ouName}`;
  try {
    const orgRes = await admin.orgunits.list({ customerId: 'my_customer' });
    const ou = orgRes.data.organizationUnits?.find(
      o => o.name === ouName || o.orgUnitPath?.endsWith(`/${ouName}`)
    );
    if (ou) orgUnitPath = ou.orgUnitPath;
  } catch (_) {}
  await admin.users.insert({
    requestBody: {
      primaryEmail,
      password,
      name: { givenName, familyName },
      orgUnitPath,
      changePasswordAtNextLogin: true
    }
  });
}

export async function moveUserToOrgUnit(tokens, userEmail, targetOuName = 'Вибувші') {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  let targetPath = `/${targetOuName}`;
  try {
    const orgRes = await admin.orgunits.list({ customerId: 'my_customer' });
    const ou = orgRes.data.organizationUnits?.find(
      o => o.name === targetOuName || o.orgUnitPath?.endsWith(`/${targetOuName}`)
    );
    if (ou) targetPath = ou.orgUnitPath;
  } catch (_) {}
  await admin.users.update({
    userKey: userEmail,
    requestBody: { orgUnitPath: targetPath }
  });
}

export async function updateUser(tokens, userEmail, { givenName, familyName, password }) {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  const body = {};
  if (givenName !== undefined || familyName !== undefined) {
    body.name = { givenName: givenName ?? '', familyName: familyName ?? '' };
  }
  if (password && password.length >= 8) body.password = password;
  if (Object.keys(body).length === 0) return;
  await admin.users.update({
    userKey: userEmail,
    requestBody: body
  });
}

export async function moveUserToOrgPath(tokens, userEmail, targetOrgUnitPath) {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  await admin.users.update({
    userKey: userEmail,
    requestBody: { orgUnitPath: targetOrgUnitPath || '/' }
  });
}

export async function searchUsersInDomain(tokens, searchTerm) {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  const users = [];
  const parts = (searchTerm || '').trim().split(/\s+/).filter(Boolean);
  const term = parts[0] || '';
  if (!term) return users;
  let pageToken;
  do {
    const res = await admin.users.list({
      customer: 'my_customer',
      query: `familyName:${term}`,
      maxResults: 100,
      pageToken,
      viewType: 'admin_view',
      orderBy: 'familyName',
      sortOrder: 'ASCENDING'
    });
    for (const u of res.data.users || []) {
      const name = [u.name?.givenName, u.name?.familyName].filter(Boolean).join(' ');
      if (u.primaryEmail) {
        users.push({
          email: u.primaryEmail,
          name: name.trim(),
          givenName: u.name?.givenName,
          familyName: u.name?.familyName,
          orgUnitPath: u.orgUnitPath
        });
      }
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return users;
}

export async function createUserInOrgPath(tokens, { givenName, familyName, primaryEmail, password }, orgUnitPath) {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  await admin.users.insert({
    requestBody: {
      primaryEmail,
      password,
      name: { givenName, familyName },
      orgUnitPath: orgUnitPath || '/',
      changePasswordAtNextLogin: true
    }
  });
}

export async function getUsersFromOrgUnit(tokens, orgUnitPath) {
  const auth = getOAuth2Client(tokens);
  const admin = google.admin({ version: 'directory_v1', auth });
  const users = [];
  let pageToken;
  do {
    const res = await admin.users.list({
      customer: 'my_customer',
      query: `orgUnitPath='${orgUnitPath}'`,
      maxResults: 500,
      pageToken,
      viewType: 'admin_view',
      orderBy: 'familyName',
      sortOrder: 'ASCENDING'
    });
    for (const u of res.data.users || []) {
      const name = [u.name?.givenName, u.name?.familyName].filter(Boolean).join(' ');
      if (u.primaryEmail) {
        users.push({
          email: u.primaryEmail,
          name: name.trim(),
          givenName: u.name?.givenName,
          familyName: u.name?.familyName
        });
      }
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return users;
}

export async function listCourses(tokens, courseStates = ['ACTIVE']) {
  const classroom = getClassroomClient(tokens);
  const courses = [];
  let pageToken;
  do {
    const res = await classroom.courses.list({
      courseStates,
      pageSize: 100,
      pageToken
    });
    for (const c of res.data.courses || []) {
      if (c.id) courses.push({ id: c.id, name: c.name, section: c.section });
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return courses;
}

async function getProfileEmail(tokens, classroom, userId) {
  try {
    const res = await classroom.userProfiles.get({ userId });
    if (res.data?.emailAddress) return res.data.emailAddress;
  } catch (_) {}
  try {
    const auth = getOAuth2Client(tokens);
    const admin = google.admin({ version: 'directory_v1', auth });
    const userRes = await admin.users.get({ userKey: userId });
    return userRes.data?.primaryEmail || '';
  } catch (_) {
    return '';
  }
}

export async function getCourseTeachers(tokens, courseId) {
  const classroom = getClassroomClient(tokens);
  const teachers = [];
  let pageToken;
  do {
    const res = await classroom.courses.teachers.list({ courseId, pageSize: 100, pageToken });
    for (const t of res.data.teachers || []) {
      const p = t.profile || {};
      let email = p.emailAddress;
      if (!email && t.userId) email = await getProfileEmail(tokens, classroom, t.userId);
      const familyName = (p.name?.familyName || '').trim();
      const givenName = (p.name?.givenName || '').trim();
      const name = ([familyName, givenName].filter(Boolean).join(' ') || p.name?.fullName || '').trim();
      teachers.push({ email: email || '', name: name || '' });
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  if (teachers.length === 0) {
    try {
      const courseRes = await classroom.courses.get({ id: courseId });
      const ownerId = courseRes.data?.ownerId;
      if (ownerId) {
        const profileRes = await classroom.userProfiles.get({ userId: ownerId });
        const p = profileRes.data || {};
        const email = p.emailAddress || '';
        const familyName = (p.name?.familyName || '').trim();
        const givenName = (p.name?.givenName || '').trim();
        const name = ([familyName, givenName].filter(Boolean).join(' ') || p.name?.fullName || '').trim();
        teachers.push({ email, name: name || '' });
      }
    } catch (_) {}
  }
  return teachers;
}

export async function getCourseStudents(tokens, courseId) {
  const classroom = getClassroomClient(tokens);
  const students = [];
  let pageToken;
  do {
    const res = await classroom.courses.students.list({ courseId, pageSize: 100, pageToken });
    for (const s of res.data.students || []) {
      const p = s.profile || {};
      let email = p.emailAddress;
      if (!email && s.userId) email = await getProfileEmail(tokens, classroom, s.userId);
      const familyName = (p.name?.familyName || '').trim();
      const givenName = (p.name?.givenName || '').trim();
      const name = (p.name?.fullName || [familyName, givenName].filter(Boolean).join(' ')).trim();
      students.push({ email: email || '', name: name || '', familyName, givenName });
    }
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return students;
}

export async function archiveCourse(tokens, courseId) {
  const classroom = getClassroomClient(tokens);
  await classroom.courses.patch({
    id: courseId,
    updateMask: 'courseState',
    requestBody: { courseState: 'ARCHIVED' }
  });
  return { success: true };
}

export async function updateCourse(tokens, courseId, name, section) {
  const classroom = getClassroomClient(tokens);
  const requestBody = {};
  if (name !== undefined) requestBody.name = name;
  if (section !== undefined) requestBody.section = section;
  if (Object.keys(requestBody).length === 0) return { success: true };
  const updateMask = Object.keys(requestBody).join(',');
  await classroom.courses.patch({
    id: courseId,
    updateMask,
    requestBody
  });
  return { success: true };
}

export async function createCourse(classroom, name, ownerId = 'me', section) {
  const res = await classroom.courses.create({
    requestBody: {
      name,
      ownerId,
      section: section || name
    }
  });
  return res.data;
}

export async function addTeacher(classroom, courseId, email) {
  try {
    await classroom.courses.teachers.create({
      courseId,
      requestBody: { userId: email }
    });
    return { success: true };
  } catch (err) {
    if (err.code === 409 || err.message?.includes('ALREADY_EXISTS')) {
      return { success: true, alreadyExists: true };
    }
    throw err;
  }
}

export async function removeTeacherFromCourse(tokens, courseId, email) {
  const classroom = getClassroomClient(tokens);
  await classroom.courses.teachers.delete({ courseId, userId: email });
  return { success: true };
}

export async function addStudentToCourse(tokens, courseId, email) {
  const classroom = getClassroomClient(tokens);
  try {
    await classroom.courses.students.create({
      courseId,
      requestBody: { userId: email }
    });
    return { success: true };
  } catch (err) {
    if (err.code === 409 || err.message?.includes('ALREADY_EXISTS')) {
      return { success: true, alreadyExists: true };
    }
    throw err;
  }
}

export async function removeStudentFromCourse(tokens, courseId, email) {
  const classroom = getClassroomClient(tokens);
  await classroom.courses.students.delete({ courseId, userId: email });
  return { success: true };
}

export function getAuthUrl() {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI || 'http://localhost:5173/callback'
  );
  return oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: [
      'https://www.googleapis.com/auth/classroom.courses',
      'https://www.googleapis.com/auth/classroom.rosters',
      'https://www.googleapis.com/auth/classroom.profile.emails',
      'https://www.googleapis.com/auth/admin.directory.user',
      'https://www.googleapis.com/auth/admin.directory.orgunit.readonly',
      'https://www.googleapis.com/auth/userinfo.email',
      'https://www.googleapis.com/auth/userinfo.profile'
    ],
    prompt: 'consent'
  });
}

export async function getUserInfo(tokens) {
  const auth = getOAuth2Client(tokens);
  const oauth2 = google.oauth2({ version: 'v2', auth });
  const res = await oauth2.userinfo.get();
  return res.data;
}

export function getTokensFromCode(code) {
  const oauth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    process.env.GOOGLE_REDIRECT_URI || 'http://localhost:5173/callback'
  );
  return oauth2Client.getToken(code);
}
