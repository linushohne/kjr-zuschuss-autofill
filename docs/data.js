import * as XLSX from './vendor/xlsx.mjs';
import { unzipSync } from './vendor/fflate.js';

export function normalize(value) {
  return String(value ?? '').normalize('NFKD').replace(/\p{M}/gu, '').toLowerCase().replace(/ß/g, 'ss').replace(/[^\p{L}\p{N}]/gu, '');
}

export function validDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value ?? '')) return false;
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return year >= 1900 && date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function isoDate(year, month, day) {
  const value = `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  return validDate(value) ? value : '';
}

export function parseBirthday(value, date1904 = false) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    const parts = XLSX.SSF.parse_date_code(value, { date1904 });
    return parts ? isoDate(parts.y, parts.m, parts.d) : '';
  }
  const text = String(value ?? '').trim();
  if (validDate(text)) return text;
  const match = /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/.exec(text);
  return match ? isoDate(Number(match[3]), Number(match[2]), Number(match[1])) : '';
}

export function ageAt(birthday, start) {
  if (!validDate(birthday) || !validDate(start)) return null;
  const [birthYear, birthMonth, birthDay] = birthday.split('-').map(Number);
  const [startYear, startMonth, startDay] = start.split('-').map(Number);
  return startYear - birthYear - Number(startMonth < birthMonth || (startMonth === birthMonth && startDay < birthDay));
}

export function eligibleForList(person, start, kind = 'participants') {
  return kind === 'participants' || (kind === 'supervisors' && ageAt(person.birthday, start) >= 18);
}

export function formatDate(value) {
  return validDate(value) ? value.split('-').reverse().join('.') : '';
}

export function formatAddress(value) {
  return String(value ?? '').replace(/\r\n?/g, '\n').trim()
    .replace(/(^|[\s,;]+)(\d{5})[ \t]+(?=\p{L})/gu, (match, separator, postcode, offset) => `${offset === 0 ? '' : '\n'}${postcode} `);
}

export function isActive(value) {
  return !['nein', 'n', 'no', 'false', '0', 'inaktiv', 'ausgetreten', 'ehemalig'].includes(normalize(value));
}

const aliases = {
  last: ['nachname', 'name', 'familienname'],
  first: ['vorname'],
  birthday: ['geburtstag', 'geburtsdatum', 'geburtsdat'],
  address: ['adresse', 'anschrift', 'wohnadresse'],
  group: ['sippe', 'gruppe'],
  level: ['stufe'],
  active: ['aktiv', 'aktivstatus'],
};

export function parseSheet(sheet, sheetName, date1904 = false) {
  if (!sheet['!ref']) throw new Error('Das Tabellenblatt ist leer.');
  const range = XLSX.utils.decode_range(sheet['!ref']);
  if (range.e.r > 5000 || range.e.c > 100) throw new Error('Das Tabellenblatt ist zu groß (max. 5.000 Zeilen und 100 Spalten).');
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true, blankrows: true });
  let columns;
  const headerIndex = rows.findIndex((row, index) => {
    if (index > 9) return false;
    const headers = row.map(normalize);
    const candidate = Object.fromEntries(Object.entries(aliases).map(([key, names]) => [key, headers.findIndex(header => names.includes(header))]));
    if (candidate.first < 0 || candidate.last < 0) return false;
    columns = candidate;
    return true;
  });
  if (headerIndex < 0) throw new Error('Spalten „Vorname“ und „Nachname“ wurden nicht gefunden.');
  const people = [];
  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const text = key => String(row[columns[key]] ?? '').trim();
    const first = text('first');
    const last = text('last');
    if (!first && !last) continue;
    people.push({
      id: `${sheetName}:${index}`,
      first,
      last,
      birthday: parseBirthday(row[columns.birthday], date1904),
      address: formatAddress(text('address')),
      group: text('group'),
      level: text('level'),
      active: isActive(row[columns.active]),
    });
  }
  return people.sort(comparePeople);
}

export function comparePeople(left, right) {
  return left.last.localeCompare(right.last, 'de') || left.first.localeCompare(right.first, 'de');
}

export function readWorkbook(buffer) {
  const bytes = new Uint8Array(buffer);
  if (bytes.length > 10 * 1024 * 1024) throw new Error('Bitte eine Datei unter 10 MB auswählen.');
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    let totalSize = 0;
    unzipSync(bytes, { filter: file => {
      totalSize += file.originalSize;
      if (totalSize > 50 * 1024 * 1024) throw new Error('Die entpackte Excel-Datei ist zu groß.');
      return false;
    } });
  }
  const workbook = XLSX.read(bytes, { type: 'array', cellDates: false, cellHTML: false, cellFormula: false, bookVBA: false });
  if (workbook.SheetNames.length > 100) throw new Error('Die Datei enthält zu viele Tabellenblätter.');
  const name = workbook.SheetNames.find(name => name.trim().toLowerCase() === 'stamm');
  if (!name) throw new Error('Das Tabellenblatt „Stamm“ fehlt.');
  return [{ name, people: parseSheet(workbook.Sheets[name], name, Boolean(workbook.Workbook?.WBProps?.date1904)) }];
}

export function personIssues(person, start) {
  const issues = [];
  if (!person.first || !person.last) issues.push('Name unvollständig');
  if (!validDate(person.birthday)) issues.push('Geburtstag fehlt oder ist ungültig');
  else if (validDate(start) && ageAt(person.birthday, start) < 0) issues.push('Geburtstag liegt nach Fahrtbeginn');
  if (!person.address) issues.push('Adresse fehlt');
  return issues;
}

export function duplicatePeople(people) {
  const seen = new Set();
  return people.some(person => {
    const key = `${normalize(person.first)}|${normalize(person.last)}|${person.birthday}|${normalize(person.address)}`;
    if (seen.has(key)) return true;
    seen.add(key);
    return false;
  });
}
