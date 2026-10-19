import { mkdir, copyFile, writeFile } from 'node:fs/promises';

await mkdir('docs/vendor', { recursive: true });
for (const [source, target] of [
  ['https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs', 'xlsx.mjs'],
  ['https://cdn.sheetjs.com/xlsx-0.20.3/package/LICENSE', 'SHEETJS-LICENSE.txt'],
]) {
  const response = await fetch(source);
  if (!response.ok) throw new Error(`Download fehlgeschlagen: ${source}`);
  await writeFile(`docs/vendor/${target}`, new Uint8Array(await response.arrayBuffer()));
}
await copyFile('node_modules/fflate/esm/browser.js', 'docs/vendor/fflate.js');
await copyFile('node_modules/fflate/LICENSE', 'docs/vendor/FFLATE-LICENSE.txt');
