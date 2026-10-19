import { unzipSync, zipSync, strFromU8, strToU8 } from './vendor/fflate.js';
import { ageAt, comparePeople, duplicatePeople, eligibleForList, formatAddress, formatDate, personIssues, validDate } from './data.js';

const wordNamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xmlNamespace = 'http://www.w3.org/XML/1998/namespace';

function descendants(element, name) {
  return Array.from(element.getElementsByTagNameNS(wordNamespace, name));
}

function children(element, name) {
  return Array.from(element.childNodes).filter(child => child.nodeType === 1 && child.namespaceURI === wordNamespace && child.localName === name);
}

function addProperty(element, name) {
  const property = element.ownerDocument.createElementNS(wordNamespace, `w:${name}`);
  element.appendChild(property);
  return property;
}

function setField(container, value, fieldName) {
  const field = descendants(container, 'ffData')[0];
  if (!field) throw new Error('Die Word-Vorlage enthält nicht die erwarteten Formularfelder.');
  const name = descendants(field, 'name')[0];
  name.setAttributeNS(wordNamespace, 'w:val', fieldName);
  const input = descendants(field, 'textInput')[0];
  for (const limit of descendants(input, 'maxLength')) limit.parentNode.removeChild(limit);
  let defaultValue = descendants(input, 'default')[0];
  if (!defaultValue) defaultValue = addProperty(input, 'default');
  defaultValue.setAttributeNS(wordNamespace, 'w:val', value);
  const paragraph = field.parentNode.parentNode.parentNode;
  let inResult = false;
  let firstRun;
  let endRun;
  for (const run of children(paragraph, 'r')) {
    const marker = descendants(run, 'fldChar')[0]?.getAttributeNS(wordNamespace, 'fldCharType');
    if (marker === 'end') {
      endRun = run;
      break;
    }
    if (inResult) {
      firstRun ??= run;
      for (const node of [...descendants(run, 't'), ...descendants(run, 'br'), ...descendants(run, 'tab')]) node.parentNode.removeChild(node);
    }
    if (marker === 'separate') inResult = true;
  }
  if (!endRun || !inResult) throw new Error('Ein Formularfeld der Word-Vorlage ist beschädigt.');
  if (!firstRun) {
    firstRun = paragraph.ownerDocument.createElementNS(wordNamespace, 'w:r');
    const properties = children(endRun, 'rPr')[0];
    if (properties) firstRun.appendChild(properties.cloneNode(true));
    paragraph.insertBefore(firstRun, endRun);
  }
  const lines = value ? value.split(/\r?\n/) : ['\u2002\u2002\u2002\u2002\u2002'];
  for (const [index, line] of lines.entries()) {
    if (index) addProperty(firstRun, 'br');
    const text = addProperty(firstRun, 't');
    text.setAttributeNS(xmlNamespace, 'xml:space', 'preserve');
    text.textContent = line;
  }
}

function prepareTable(table, people, start, rowCount, supervisors = false) {
  const rows = children(table, 'tr');
  const header = rows[0];
  const sample = rows[1].cloneNode(true);
  let headerProperties = children(header, 'trPr')[0];
  if (!headerProperties) headerProperties = header.insertBefore(table.ownerDocument.createElementNS(wordNamespace, 'w:trPr'), header.firstChild);
  addProperty(headerProperties, 'tblHeader');
  for (const row of rows.slice(1)) table.removeChild(row);
  for (let index = 0; index < rowCount; index += 1) {
    const row = sample.cloneNode(true);
    let properties = children(row, 'trPr')[0];
    if (!properties) properties = row.insertBefore(table.ownerDocument.createElementNS(wordNamespace, 'w:trPr'), row.firstChild);
    addProperty(properties, 'cantSplit');
    for (const height of children(properties, 'trHeight')) height.setAttributeNS(wordNamespace, 'w:hRule', 'atLeast');
    const person = people[index];
    const number = index + 1;
    const values = person
      ? supervisors
        ? [String(number), `${person.last}, ${person.first}`, formatAddress(person.address), '', '', String(ageAt(person.birthday, start)), '']
        : [String(number), `${person.last}, ${person.first}`, formatAddress(person.address), 'X', '', '', String(ageAt(person.birthday, start)), '']
      : Array(supervisors ? 7 : 8).fill('');
    for (const [cellIndex, cell] of children(row, 'tc').entries()) {
      const paragraphs = descendants(cell, 'p').filter(paragraph => descendants(paragraph, 'ffData').length);
      for (const [fieldIndex, paragraph] of paragraphs.entries()) {
        if (supervisors && paragraphs.length > 1) {
          for (const spacing of descendants(paragraph, 'spacing')) {
            spacing.setAttributeNS(wordNamespace, 'w:before', '0');
            spacing.setAttributeNS(wordNamespace, 'w:after', '0');
          }
        }
        setField(paragraph, fieldIndex === 0 ? values[cellIndex] : '', `person_${number}_${cellIndex}${fieldIndex ? `_${fieldIndex}` : ''}`);
      }
    }
    table.appendChild(row);
  }
}

export function createWord(template, people, trip, xml = { DOMParser, XMLSerializer }) {
  const kind = trip.kind || 'participants';
  if (!['participants', 'supervisors'].includes(kind)) throw new Error('Unbekannte Listenart.');
  if (!validDate(trip.start)) throw new Error('Bitte einen gültigen Fahrtbeginn eintragen.');
  if (trip.end && (!validDate(trip.end) || trip.end < trip.start)) throw new Error('Das Fahrtende darf nicht vor dem Fahrtbeginn liegen.');
  if (!people.length) throw new Error('Bitte mindestens eine Person auswählen.');
  if (people.some(person => personIssues(person, trip.start).length)) throw new Error('Bitte die fehlenden oder ungültigen Angaben der ausgewählten Personen ergänzen.');
  if (people.some(person => !eligibleForList(person, trip.start, kind))) throw new Error('Betreuer müssen zum Fahrtbeginn mindestens 18 Jahre alt sein.');
  if (duplicatePeople(people)) throw new Error('Eine Person ist doppelt ausgewählt. Bitte die Auswahl prüfen.');
  const selected = [...people].sort(comparePeople);
  const files = unzipSync(new Uint8Array(template));
  const document = new xml.DOMParser().parseFromString(strFromU8(files['word/document.xml']), 'application/xml');
  const body = descendants(document, 'body')[0];
  const tables = children(body, 'tbl');
  const supervisors = kind === 'supervisors';
  if (tables.length !== (supervisors ? 2 : 5)) throw new Error('Die Word-Vorlage hat ein unerwartetes Format.');
  const tripFields = descendants(tables[supervisors ? 0 : 1], 'tc').filter(cell => descendants(cell, 'ffData').length);
  const dates = trip.end ? `${formatDate(trip.start)} – ${formatDate(trip.end)}` : `${formatDate(trip.start)} –`;
  for (const [index, value] of [trip.title || '', trip.location || '', dates].entries()) {
    setField(tripFields[index], value, ['fahrt_titel', 'fahrt_ort', 'fahrt_zeitraum'][index]);
  }
  const firstTable = tables[supervisors ? 1 : 3];
  const firstRows = supervisors ? 5 : children(firstTable, 'tr').length - 1;
  const continuationRows = supervisors ? 10 : children(tables[4], 'tr').length - 1;
  const rowCount = firstRows + Math.max(1, Math.ceil(Math.max(0, selected.length - firstRows) / continuationRows)) * continuationRows;
  let next = firstTable.nextSibling;
  while (!supervisors && next && !(next.nodeType === 1 && next.localName === 'sectPr')) {
    const following = next.nextSibling;
    body.removeChild(next);
    next = following;
  }
  prepareTable(firstTable, selected, trip.start, rowCount, supervisors);
  const trailing = document.createElementNS(wordNamespace, 'w:p');
  const trailingProperties = addProperty(trailing, 'pPr');
  const trailingSpacing = addProperty(trailingProperties, 'spacing');
  trailingSpacing.setAttributeNS(wordNamespace, 'w:line', '20');
  trailingSpacing.setAttributeNS(wordNamespace, 'w:lineRule', 'exact');
  trailingSpacing.setAttributeNS(wordNamespace, 'w:after', '0');
  body.insertBefore(trailing, next);
  files['word/document.xml'] = strToU8(new xml.XMLSerializer().serializeToString(document));
  return zipSync(files);
}
