import { mkdir, writeFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
import * as XLSX from '../docs/vendor/xlsx.mjs';

export function fixtureWorkbook() {
  const workbook = XLSX.utils.book_new();
  const rows = [
    ['Nachname', 'Vorname', 'Aktiv', 'Geburtstag', 'Adresse', 'Sippe', 'Stufe', 'Allergien'],
    ['Muster', 'Anna', 'Ja', '20.10.2014', 'Beispielstraße 12\n80331 München', 'Füchse', 'Jungpfadfinder', 'NICHT_EXPORTIEREN'],
    ['Beispiel', 'Ben', 'Ja', 42064, 'Testweg 3, 80333 München', 'Eulen', 'Wölflinge', 'NICHT_EXPORTIEREN'],
    ['Alt', 'Carla', 'Nein', '01.01.2010', 'Musterweg 4, 80331 München', 'Füchse', 'Pfadfinder', 'NICHT_EXPORTIEREN'],
    ['Fehlend', 'Dana', 'Ja', '', '', 'Eulen', 'Wölflinge', 'NICHT_EXPORTIEREN'],
    ['<script>', 'Änne & Ömer', 'Ja', '29.02.2012', 'Am <Test> & Weg 5, 80331 München', 'Füchse', 'Jungpfadfinder', 'NICHT_EXPORTIEREN'],
  ];
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Stamm');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['Vorname', 'Adresse', 'Nachname', 'Geburtstag', 'Sippe'],
    ['Emil', 'Gruppenweg 1, 80331 München', 'Gruppe', '2015-03-01', 'Füchse'],
  ]), 'Gruppe');
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([['Notizen'], ['Keine Mitgliederliste']]), 'Notizen');
  return XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
}

export function fixturePeople(count) {
  return Array.from({ length: count }, (_, index) => ({
    id: String(index),
    first: index === 0 ? 'Änne & Ömer' : `Vorname ${index + 1}`,
    last: `Beispiel ${String(index + 1).padStart(2, '0')}`,
    birthday: '2014-10-20',
    address: index === 0 ? 'Langer Beispielstraßenname 123 a\nHinterhaus, 2. Stock links\n80331 München' : 'Musterstraße 12\n80331 München',
    group: 'Füchse', level: 'Jungpfadfinder', active: true,
  }));
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  await mkdir('/tmp/kjr-qa', { recursive: true });
  await writeFile('/tmp/kjr-qa/Beispieldaten.xlsx', new Uint8Array(fixtureWorkbook()));
}
