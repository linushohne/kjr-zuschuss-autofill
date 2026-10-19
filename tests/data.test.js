import test from 'node:test';
import assert from 'node:assert/strict';
import * as XLSX from '../docs/vendor/xlsx.mjs';
import { ageAt, duplicatePeople, isActive, normalize, parseBirthday, parseSheet, personIssues, readWorkbook, validDate } from '../docs/data.js';
import { fixturePeople, fixtureWorkbook } from '../scripts/qa-fixtures.mjs';

test('Alter vor, am und nach Geburtstag', () => {
  assert.equal(ageAt('2014-10-20', '2026-10-19'), 11);
  assert.equal(ageAt('2014-10-20', '2026-10-20'), 12);
  assert.equal(ageAt('2014-10-20', '2026-10-21'), 12);
  assert.equal(ageAt('2027-01-01', '2026-10-19'), -1);
  assert.equal(ageAt('', '2026-10-19'), null);
});

test('Geburtstage mit Schaltjahr und ungültige Datumswerte', () => {
  assert.equal(parseBirthday('29.02.2012'), '2012-02-29');
  assert.equal(parseBirthday('29.02.2013'), '');
  assert.equal(parseBirthday('1/3/2015'), '2015-03-01');
  assert.equal(parseBirthday('2015-03-01'), '2015-03-01');
  assert.equal(parseBirthday('01.01.15'), '');
  assert.equal(validDate('2026-02-30'), false);
  assert.equal(validDate('2026-13-01'), false);
  assert.equal(validDate('2026-10-19'), true);
  assert.equal(ageAt('2012-02-29', '2026-02-28'), 13);
  assert.equal(ageAt('2012-02-29', '2026-03-01'), 14);
});

test('Excel-Datumssysteme 1900 und 1904', () => {
  assert.equal(parseBirthday(42064), '2015-03-01');
  assert.equal(parseBirthday(40602, true), '2015-03-01');
  assert.equal(parseBirthday(60), '');
});

test('Aktiv-Status erkennt explizite Inaktiv-Werte', () => {
  for (const value of ['Nein', false, 0, 'inaktiv']) assert.equal(isActive(value), false);
  for (const value of ['Ja', true, 1, 'X', '']) assert.equal(isActive(value), true);
});

test('Suche erkennt Umlaute und erhält nichtlateinische Namen', () => {
  assert.equal(normalize('Müller Straße'), 'mullerstrasse');
  assert.equal(normalize('李'), '李');
});

test('Klassische XLS-Dateien werden ebenfalls eingelesen', () => {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['Nachname', 'Vorname', 'Geburtstag', 'Adresse'],
    ['Müller', 'Änne', '01.03.2015', 'Testweg 1'],
  ]), 'Stamm');
  const sheets = readWorkbook(XLSX.write(workbook, { type: 'array', bookType: 'biff8' }));
  assert.equal(sheets[0].people[0].first, 'Änne');
  assert.equal(sheets[0].people[0].last, 'Müller');
});

test('Nur benötigte Daten werden aus der Arbeitsmappe übernommen', () => {
  const sheets = readWorkbook(fixtureWorkbook());
  assert.equal(sheets[0].name, 'Stamm');
  assert.equal(sheets[0].people.length, 5);
  assert.equal(sheets[1].people[0].address, 'Gruppenweg 1, 80331 München');
  assert.match(sheets[2].error, /Spalten/);
  assert.equal(JSON.stringify(sheets).includes('NICHT_EXPORTIEREN'), false);
  assert.equal(sheets[0].people.find(person => person.first === 'Carla').active, false);
});

test('Header können nach einer Titelzeile stehen', () => {
  const sheet = XLSX.utils.aoa_to_sheet([['Titel'], [], ['Vorname', 'Name', 'Geburtsdatum', 'Anschrift'], ['Anna', 'Muster', '1.3.2015', 'Testweg 1']]);
  assert.equal(parseSheet(sheet, 'Test')[0].birthday, '2015-03-01');
});

test('Fehlende Angaben und Geburtstag nach Fahrtbeginn werden erkannt', () => {
  const person = fixturePeople(1)[0];
  assert.deepEqual(personIssues(person, '2026-10-19'), []);
  assert.equal(personIssues({ ...person, birthday: '', address: '' }, '2026-10-19').length, 2);
  assert.match(personIssues({ ...person, birthday: '2027-01-01' }, '2026-10-19')[0], /nach Fahrtbeginn/);
});

test('Doppelte Personen werden erkannt, nicht still entfernt', () => {
  const person = fixturePeople(1)[0];
  assert.equal(duplicatePeople([person, { ...person, id: 'other' }]), true);
  assert.equal(duplicatePeople([person, { ...person, birthday: '2014-10-21' }]), false);
});

test('Unpassende und zu große Dateien werden abgewiesen', () => {
  assert.throws(() => readWorkbook(new Uint8Array(11 * 1024 * 1024)), /10 MB/);
  assert.throws(() => readWorkbook(new Uint8Array([1, 2, 3])), /./);
  assert.throws(() => parseSheet({ '!ref': 'A1:ZZ10000' }, 'Groß'), /zu groß/);
});
