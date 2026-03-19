import * as cheerio from 'cheerio';

export function parseJournalList(html) {
  const $ = cheerio.load(html);
  const items = [];
  const tableSelector = 'table.journal-choose tbody tr';
  let rows = $(tableSelector);
  if (!rows.length) {
    rows = $('table tbody tr').filter((_, r) => $(r).find('a[href*="journal="]').length > 0);
  }
  rows.each((_, row) => {
    const tds = $(row).find('td');
    if (tds.length < 2) return;
    const subject = $(tds[0]).text().trim();
    const links = [];
    $(tds[1]).find('a').each((_, a) => {
      const href = $(a).attr('href') || '';
      const text = $(a).text().trim();
      const match = href.match(/journal=(\d+)(?:&subgroup=(\d+))?/);
      if (match) {
        links.push({
          journalId: match[1],
          subgroupId: match[2] || null,
          classLabel: text,
          href
        });
      }
    });
    if (subject && links.length) {
      items.push({ subject, classes: links });
    }
  });
  return items;
}

export function parseJournalPage(html) {
  const $ = cheerio.load(html);
  const teacherLink = $('a.teacher__link').first();
  if (!teacherLink.length) return null;
  let name = teacherLink.text().trim();
  const parts = name.split(/\s+/).filter(Boolean);
  if (parts.length >= 3) name = parts.slice(0, -1).join(' ');
  const href = teacherLink.attr('href') || '';
  const profileId = href.match(/\/id(\d+)/)?.[1];
  return { name, profileId, profileUrl: href ? `https://nz.ua${href}` : null };
}

export function parseJournalStudents(html) {
  const $ = cheerio.load(html);
  const students = [];
  $('#journalList tbody tr').each((_, row) => {
    const td = $(row).find('td.pt-theme');
    if (!td.length) return;
    const link = td.find('a').first();
    const name = (link.length ? link.text() : td.text()).trim();
    const href = link.attr('href') || '';
    const studentId = td.attr('data-student-id') || href.match(/\/id(\d+)/)?.[1];
    if (name) students.push({ name, studentId, profileUrl: href ? `https://nz.ua${href}` : null });
  });
  return students;
}

export function parseJournalDetails(html) {
  const teacher = parseJournalPage(html);
  const students = parseJournalStudents(html);
  return { teacher, students };
}
