import { readFile, writeFile } from 'node:fs/promises';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { unzipSync, zipSync, strFromU8, strToU8 } from '../docs/vendor/fflate.js';

const input = process.argv[2];
if (!input) throw new Error('Aufruf: node scripts/prepare-template.mjs /pfad/zur/leeren-vorlage.docx');
const files = unzipSync(new Uint8Array(await readFile(input)));
const parser = new DOMParser();
const serializer = new XMLSerializer();
const relationships = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
for (const [name, bytes] of Object.entries(files)) {
  if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
  const document = parser.parseFromString(strFromU8(bytes), 'application/xml');
  if (name.endsWith('.rels')) {
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      if (element.getAttribute('TargetMode') === 'External') element.parentNode.removeChild(element);
    }
  }
  if (name === 'word/document.xml') {
    for (const element of Array.from(document.getElementsByTagName('*'))) {
      element.removeAttributeNS(relationships, 'link');
    }
  }
  if (name.startsWith('docProps/')) {
    const root = document.documentElement;
    while (root.firstChild) root.removeChild(root.firstChild);
  }
  files[name] = strToU8(serializer.serializeToString(document));
}
await writeFile('docs/teilnehmerliste.docx', zipSync(files));
