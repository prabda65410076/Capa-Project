// Builds the web page from src/: node build.mjs  ->  index.html (one file, works offline)
//   node build.mjs --artifact <path>   also writes the page without the <html>/<head> wrapper
//                                      (for publishing as a claude.ai artifact)
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const root = dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const { coreSource } = require('./tools/core.js');

const read = (p) => readFileSync(join(root, p), 'utf8');
const noScriptEnd = (s) => s.replace(/<\/script/gi, '<\\/script');

const page = read('src/ui/page.html');
const [, headAndBody] = page.split('<!--HEAD-->');
const [head0, body0] = headAndBody.split('<!--BODY-->');
const sample = readFileSync(join(root, 'src/ui/sample.xlsx')).toString('base64');

const head = head0.split('/*STYLE*/').join(read('src/ui/style.css')).trim();
const body = body0
  .split('/*SAMPLE*/').join(sample)
  .split('/*CORE*/').join(noScriptEnd(coreSource()))
  .split('/*APP*/').join(noScriptEnd(read('src/ui/app.js')))
  .trim();

const html = '<!doctype html>\n<html lang="th">\n<head>\n<meta charset="utf-8">\n' +
  '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n' +
  head + '\n</head>\n<body>\n' + body + '\n</body>\n</html>\n';
writeFileSync(join(root, 'index.html'), html);
console.log('index.html', (html.length / 1024).toFixed(0) + ' KB');

const i = process.argv.indexOf('--artifact');
if (i > 0 && process.argv[i + 1]) {
  const out = process.argv[i + 1];
  mkdirSync(dirname(out), { recursive: true });
  writeFileSync(out, head + '\n' + body + '\n');
  console.log(out);
}
