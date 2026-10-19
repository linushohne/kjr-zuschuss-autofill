import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzipSync, strFromU8 } from '../docs/vendor/fflate.js';
import { createWord } from '../docs/word.js';
import { fixturePeople } from '../scripts/qa-fixtures.mjs';

const xml = { DOMParser, XMLSerializer };
const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const template = new Uint8Array(await readFile(new URL('../docs/teilnehmerliste.docx', import.meta.url)));
const supervisorTemplate = new Uint8Array(await readFile(new URL('../docs/betreuerliste.docx', import.meta.url)));
const trip = { start: '2026-10-19', title: '', location: '', end: '' };
const descendants = (node, name) => Array.from(node.getElementsByTagNameNS(namespace, name));

for (const count of [1, 5, 6, 14, 15, 16, 25, 26]) {
  test(`Word mit ${count} Personen: Formularfelder und Folgeseiten`, async () => {
    const output = createWord(template, fixturePeople(count), trip, xml);
    const files = unzipSync(output);
    const document = new DOMParser().parseFromString(strFromU8(files['word/document.xml']), 'application/xml');
    assert.equal(descendants(document, 'tbl').length, 4);
    const rowCount = 5 + Math.max(1, Math.ceil(Math.max(0, count - 5) / 10)) * 10;
    const rows = descendants(descendants(document, 'tbl')[3], 'tr');
    assert.equal(rows.length, rowCount + 1);
    const fields = descendants(document, 'ffData');
    const names = fields.map(field => descendants(field, 'name')[0].getAttributeNS(namespace, 'val'));
    assert.equal(new Set(names).size, names.length);
    assert.equal(descendants(document, 'instrText').length, fields.length);
    const fieldValue = name => {
      const field = fields.find(field => descendants(field, 'name')[0].getAttributeNS(namespace, 'val') === name);
      return descendants(field, 'default')[0].getAttributeNS(namespace, 'val');
    };
    assert.equal(fieldValue('fahrt_titel'), '');
    assert.equal(fieldValue('fahrt_zeitraum'), '19.10.2026 –');
    assert.equal(fieldValue('person_1_1'), 'Beispiel 01, Änne & Ömer');
    assert.equal(fieldValue('person_1_3'), 'X');
    assert.equal(fieldValue('person_1_4'), '');
    assert.equal(fieldValue('person_1_5'), '');
    assert.equal(fieldValue('person_1_6'), '11');
    assert.equal(fieldValue('person_1_7'), '');
    assert.equal(fieldValue(`person_${count}_0`), String(count));
    for (let number = count + 1; number <= rowCount; number += 1) {
      for (let column = 0; column < 8; column += 1) {
        assert.equal(fieldValue(`person_${number}_${column}`), '');
      }
      assert.equal(descendants(rows[number], 't').map(node => node.textContent).join('').trim(), '');
      assert.equal(descendants(rows[number], 'ffData').length, 8);
    }
    assert.equal(descendants(document, 'pageBreakBefore').length, 0);
    assert.equal(descendants(document, 'tblHeader').length, 1);
    for (const name of ['word/media/image1.jpeg', 'word/media/image2.jpeg', 'word/styles.xml', 'word/settings.xml']) {
      assert.deepEqual(files[name], unzipSync(template)[name]);
    }
    const displayedText = descendants(document, 't').map(node => node.textContent).join(' ');
    assert.match(displayedText, /Änne & Ömer/);
    assert.match(displayedText, /80331 München/);
    await mkdir('/tmp/kjr-qa', { recursive: true });
    await writeFile(`/tmp/kjr-qa/liste-${count}.docx`, output);
  });
}

test('Einzeilige Adressen erhalten einen echten Word-Zeilenumbruch vor der PLZ', async () => {
  const people = [{ ...fixturePeople(1)[0], address: 'Testweg 12, 01234 Beispielstadt' }];
  const output = createWord(template, people, trip, xml);
  const files = unzipSync(output);
  const document = new DOMParser().parseFromString(strFromU8(files['word/document.xml']), 'application/xml');
  const row = descendants(descendants(document, 'tbl')[3], 'tr')[1];
  const cell = descendants(row, 'tc')[2];
  assert.equal(descendants(cell, 'default')[0].getAttributeNS(namespace, 'val'), 'Testweg 12\n01234 Beispielstadt');
  assert.equal(descendants(cell, 'br').length, 1);
  assert.deepEqual(descendants(cell, 't').map(node => node.textContent), ['Testweg 12', '01234 Beispielstadt']);
  await mkdir('/tmp/kjr-qa', { recursive: true });
  await writeFile('/tmp/kjr-qa/liste-adresse.docx', output);
});

test('Optionale Fahrtdaten werden in bestehende Felder geschrieben', () => {
  const files = unzipSync(createWord(template, fixturePeople(1), { ...trip, title: 'Herbst & <Lager>', location: 'München', end: '2026-10-22' }, xml));
  const document = new DOMParser().parseFromString(strFromU8(files['word/document.xml']), 'application/xml');
  const text = descendants(document, 't').map(node => node.textContent).join(' ');
  assert.match(text, /Herbst & <Lager>/);
  assert.match(text, /19.10.2026 – 22.10.2026/);
});

test('Ungültige Exporte werden blockiert', () => {
  assert.throws(() => createWord(template, fixturePeople(1), { ...trip, start: '' }, xml), /Fahrtbeginn/);
  assert.throws(() => createWord(template, fixturePeople(1), { ...trip, end: '2026-10-18' }, xml), /Fahrtende/);
  assert.throws(() => createWord(template, [], trip, xml), /mindestens/);
  assert.throws(() => createWord(template, [{ ...fixturePeople(1)[0], birthday: '' }], trip, xml), /Angaben/);
  assert.throws(() => createWord(template, [fixturePeople(1)[0], fixturePeople(1)[0]], trip, xml), /doppelt/);
});

for (const count of [1, 6, 16]) {
  test(`Betreuerliste mit ${count} Volljährigen und leeren Zusatzfeldern`, async () => {
    const people = fixturePeople(count).map(person => ({ ...person, birthday: '2008-10-19' }));
    const output = createWord(supervisorTemplate, people, { ...trip, kind: 'supervisors', title: 'Herbst & Lager', location: 'München', end: '2026-10-22' }, xml);
    const files = unzipSync(output);
    const document = new DOMParser().parseFromString(strFromU8(files['word/document.xml']), 'application/xml');
    const tables = descendants(document, 'tbl');
    assert.equal(tables.length, 2);
    const rows = descendants(tables[1], 'tr');
    const rowCount = 5 + Math.max(1, Math.ceil(Math.max(0, count - 5) / 10)) * 10;
    assert.equal(rows.length, rowCount + 1);
    const fields = descendants(document, 'ffData');
    const names = fields.map(field => descendants(field, 'name')[0].getAttributeNS(namespace, 'val'));
    assert.equal(new Set(names).size, names.length);
    const fieldValue = name => descendants(fields.find(field => descendants(field, 'name')[0].getAttributeNS(namespace, 'val') === name), 'default')[0].getAttributeNS(namespace, 'val');
    assert.equal(fieldValue('fahrt_titel'), 'Herbst & Lager');
    assert.equal(fieldValue('fahrt_ort'), 'München');
    assert.equal(fieldValue('fahrt_zeitraum'), '19.10.2026 – 22.10.2026');
    assert.equal(fieldValue('person_1_1'), 'Beispiel 01, Änne & Ömer');
    assert.equal(fieldValue(`person_${count}_0`), String(count));
    for (let number = 1; number <= rowCount; number += 1) {
      assert.equal(fieldValue(`person_${number}_3`), '');
      assert.equal(fieldValue(`person_${number}_3_1`), '');
      assert.equal(fieldValue(`person_${number}_4`), '');
      assert.equal(fieldValue(`person_${number}_5`), number <= count ? '18' : '');
      assert.equal(descendants(descendants(rows[number], 'tc')[6], 't').map(node => node.textContent).join('').trim(), '');
      if (number > count) assert.equal(descendants(rows[number], 't').map(node => node.textContent).join('').trim(), '');
    }
    assert.equal(descendants(document, 'tblHeader').length, 1);
    assert.equal(descendants(document, 'trHeight').some(node => node.getAttributeNS(namespace, 'hRule') === 'exact'), false);
    const text = descendants(document, 't').map(node => node.textContent).join(' ');
    assert.match(text, /Betreu/);
    assert.match(text, /Einsichtnahmeverfahren/);
    assert.equal(text.includes('X'), false);
    await mkdir('/tmp/kjr-qa', { recursive: true });
    await writeFile(`/tmp/kjr-qa/betreuer-${count}.docx`, output);
  });
}

test('Minderjährige und unpassende Vorlagen werden für Betreuer blockiert', () => {
  assert.throws(() => createWord(supervisorTemplate, fixturePeople(1), { ...trip, kind: 'supervisors' }, xml), /mindestens 18/);
  const adult = { ...fixturePeople(1)[0], birthday: '2008-10-19' };
  assert.throws(() => createWord(template, [adult], { ...trip, kind: 'supervisors' }, xml), /Vorlage/);
  assert.throws(() => createWord(supervisorTemplate, [adult], trip, xml), /Vorlage/);
});

for (const [name, source] of [['Teilnehmer', template], ['Betreuer', supervisorTemplate]]) {
  test(`${name}-Vorlage enthält keine externen Beziehungen oder Autorenmetadaten`, () => {
    const files = unzipSync(source);
    for (const [name, bytes] of Object.entries(files)) {
      if (name.endsWith('.rels')) assert.equal(strFromU8(bytes).includes('TargetMode="External"'), false);
      if (name.startsWith('docProps/')) {
        const document = new DOMParser().parseFromString(strFromU8(bytes), 'application/xml');
        assert.equal(document.documentElement.childNodes.length, 0);
      }
    }
  });
}
