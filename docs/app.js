import { ageAt, comparePeople, duplicatePeople, normalize, personIssues, readWorkbook, validDate } from './data.js';
import { createWord } from './word.js';

const element = id => document.getElementById(id);
let sheets = [];
let people = [];
let selected = new Set();
let visible = [];
let editing;
let revision = 0;
let exporting = false;

function trip() {
  return Object.fromEntries(['start', 'end', 'title', 'location'].map(key => [key, element(key).value.trim()]));
}

function error(message = '') {
  element('error').textContent = message;
  element('error').hidden = !message;
}

function optionList(select, values, first) {
  select.replaceChildren();
  if (first) select.add(new Option(first, ''));
  for (const value of values) select.add(new Option(value, value));
}

function selectedPeople() {
  return people.filter(person => selected.has(person.id));
}

function updateDownload() {
  const picked = selectedPeople();
  const values = trip();
  const incomplete = picked.filter(person => personIssues(person, values.start).length).length;
  let hint = '';
  if (!validDate(values.start)) hint = 'Fahrtbeginn eintragen.';
  else if (values.end && (!validDate(values.end) || values.end < values.start)) hint = 'Fahrtende prüfen.';
  else if (!picked.length) hint = 'Personen auswählen.';
  else if (incomplete) hint = `${incomplete} ${incomplete === 1 ? 'Person: Angaben prüfen.' : 'Personen: Angaben prüfen.'}`;
  else if (duplicatePeople(picked)) hint = 'Doppelte Person in der Auswahl.';
  element('selection').textContent = `${picked.length} ausgewählt · ${visible.length} angezeigt`;
  element('clear-selection').hidden = picked.length === 0;
  element('download-hint').textContent = hint;
  element('download').disabled = Boolean(hint) || exporting;
  element('download').textContent = exporting ? 'Wird erstellt …' : 'Word herunterladen';
  const pickedVisible = visible.filter(person => selected.has(person.id)).length;
  element('select-all').checked = visible.length > 0 && pickedVisible === visible.length;
  element('select-all').indeterminate = pickedVisible > 0 && pickedVisible < visible.length;
  element('select-all').disabled = visible.length === 0;
}

function openEditor(person) {
  editing = person;
  for (const key of ['first', 'last', 'birthday', 'address']) element(`edit-${key}`).value = person[key];
  element('edit-birthday').max = element('start').value || '';
  element('edit-dialog').showModal();
}

function render() {
  const query = normalize(element('search').value);
  visible = people.filter(person =>
    (!element('active').checked || person.active) &&
    (!element('group').value || person.group === element('group').value) &&
    (!element('level').value || person.level === element('level').value) &&
    (!query || normalize(`${person.first} ${person.last} ${person.last} ${person.first} ${person.address}`).includes(query))
  );
  const fragment = document.createDocumentFragment();
  for (const person of visible) {
    const row = document.createElement('tr');
    const selectCell = row.insertCell();
    selectCell.className = 'select-cell';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = selected.has(person.id);
    checkbox.setAttribute('aria-label', `${person.first} ${person.last} auswählen`);
    checkbox.addEventListener('change', () => {
      if (checkbox.checked) selected.add(person.id);
      else selected.delete(person.id);
      error();
      updateDownload();
    });
    selectCell.appendChild(checkbox);
    const nameCell = row.insertCell();
    nameCell.className = 'person-name';
    nameCell.textContent = `${person.last}, ${person.first}`;
    const issues = personIssues(person, element('start').value);
    if (issues.length) {
      const warning = document.createElement('span');
      warning.className = 'issue';
      warning.textContent = issues.join(' · ');
      nameCell.appendChild(warning);
    }
    row.insertCell().textContent = person.group || '–';
    const ageCell = row.insertCell();
    ageCell.className = 'age';
    const age = ageAt(person.birthday, element('start').value);
    ageCell.textContent = age !== null && age >= 0 ? String(age) : '–';
    const addressCell = row.insertCell();
    addressCell.className = 'address';
    addressCell.textContent = person.address || '–';
    const edit = document.createElement('button');
    edit.type = 'button';
    edit.className = 'text-button';
    edit.textContent = 'Bearbeiten';
    edit.setAttribute('aria-label', `${person.first} ${person.last} bearbeiten`);
    edit.addEventListener('click', () => openEditor(person));
    row.insertCell().appendChild(edit);
    fragment.appendChild(row);
  }
  element('rows').replaceChildren(fragment);
  element('empty').hidden = visible.length > 0;
  updateDownload();
}

function chooseSheet() {
  selected.clear();
  people = sheets.find(sheet => sheet.name === element('sheet').value)?.people || [];
  element('search').value = '';
  const unique = key => [...new Set(people.map(person => person[key]).filter(Boolean))].sort((left, right) => left.localeCompare(right, 'de'));
  optionList(element('group'), unique('group'), 'Alle');
  optionList(element('level'), unique('level'), 'Alle');
  const sheetError = sheets.find(sheet => sheet.name === element('sheet').value)?.error;
  error(sheetError || '');
  render();
}

function clearData() {
  sheets = [];
  people = [];
  visible = [];
  selected = new Set();
  element('file').value = '';
  element('filename').textContent = 'Keine Datei ausgewählt';
  element('people').hidden = true;
  element('reset').hidden = true;
  element('rows').replaceChildren();
  element('sheet').replaceChildren();
  element('group').replaceChildren();
  element('level').replaceChildren();
  element('status').textContent = '';
  editing = undefined;
  if (element('edit-dialog').open) element('edit-dialog').close();
  element('edit-form').reset();
  error();
}

element('file').addEventListener('change', async event => {
  const file = event.target.files[0];
  if (!file) return;
  const currentRevision = ++revision;
  clearData();
  element('status').textContent = 'Excel wird gelesen …';
  try {
    if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('Bitte eine Excel-Datei (.xlsx oder .xls) auswählen.');
    if (file.size > 10 * 1024 * 1024) throw new Error('Bitte eine Datei unter 10 MB auswählen.');
    const buffer = await file.arrayBuffer();
    if (currentRevision !== revision) return;
    sheets = readWorkbook(buffer);
    optionList(element('sheet'), sheets.map(sheet => sheet.name));
    element('sheet').value = sheets.find(sheet => normalize(sheet.name) === 'stamm' && !sheet.error)?.name || sheets.find(sheet => !sheet.error).name;
    element('filename').textContent = file.name;
    element('people').hidden = false;
    element('reset').hidden = false;
    element('status').textContent = '';
    chooseSheet();
  } catch (reason) {
    if (currentRevision !== revision) return;
    clearData();
    error(reason instanceof Error ? reason.message : 'Die Datei konnte nicht gelesen werden.');
  }
});

element('reset').addEventListener('click', () => {
  revision += 1;
  clearData();
  for (const key of ['start', 'end', 'title', 'location', 'search']) element(key).value = '';
  element('active').checked = true;
  document.querySelector('details').open = false;
});

element('sheet').addEventListener('change', chooseSheet);
for (const id of ['search', 'group', 'level', 'active', 'start']) element(id).addEventListener('input', render);
element('end').addEventListener('input', updateDownload);
element('select-all').addEventListener('change', event => {
  for (const person of visible) {
    if (event.target.checked) selected.add(person.id);
    else selected.delete(person.id);
  }
  error();
  render();
});
element('clear-selection').addEventListener('click', () => {
  selected.clear();
  error();
  render();
});
element('edit-cancel').addEventListener('click', () => element('edit-dialog').close());
element('edit-dialog').addEventListener('close', () => {
  editing = undefined;
  element('edit-form').reset();
});
element('edit-form').addEventListener('submit', event => {
  event.preventDefault();
  if (!editing) return;
  for (const key of ['first', 'last', 'birthday', 'address']) editing[key] = element(`edit-${key}`).value.trim();
  people.sort(comparePeople);
  element('edit-dialog').close();
  error();
  render();
});

element('download').addEventListener('click', async () => {
  error();
  const picked = selectedPeople().map(person => ({ ...person }));
  const values = trip();
  const currentRevision = revision;
  exporting = true;
  updateDownload();
  try {
    const response = await fetch('./teilnehmerliste.docx');
    if (!response.ok) throw new Error('Die Word-Vorlage konnte nicht geladen werden. Bitte erneut versuchen.');
    const template = await response.arrayBuffer();
    if (currentRevision !== revision) return;
    const bytes = createWord(template, picked, values);
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    const link = document.createElement('a');
    link.href = url;
    link.download = `Teilnehmerliste_${values.start}.docx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (reason) {
    if (currentRevision === revision) error(reason instanceof Error ? reason.message : 'Die Word-Datei konnte nicht erstellt werden.');
  } finally {
    exporting = false;
    updateDownload();
  }
});
