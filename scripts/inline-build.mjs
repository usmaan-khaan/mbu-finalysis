import { readFile, writeFile, readdir, unlink, rmdir } from 'node:fs/promises';
import { resolve, dirname, basename } from 'node:path';

const dist = resolve('dist');
const index = resolve(dist, 'index.html');
let html = await readFile(index, 'utf8');
const assets = [];

async function inlineAsset(reference, kind) {
  if (!/^\/?assets\/[a-zA-Z0-9._-]+$/.test(reference)) throw new Error(`Unexpected ${kind} reference: ${reference}`);
  const asset = resolve(dist, reference.replace(/^\//, ''));
  if (dirname(asset) !== resolve(dist, 'assets')) throw new Error(`Asset outside dist: ${reference}`);
  assets.push(asset);
  return await readFile(asset, 'utf8');
}

for (const match of [...html.matchAll(/<link\b[^>]*\brel="stylesheet"[^>]*\bhref="([^"]+)"[^>]*>/g)]) {
  const css = (await inlineAsset(match[1], 'stylesheet')).replace(/<\/style/gi, '<\\/style');
  html = html.replace(match[0], () => `<style>${css}</style>`);
}
for (const match of [...html.matchAll(/<script\b[^>]*\bsrc="([^"]+)"[^>]*><\/script>/g)]) {
  const javascript = (await inlineAsset(match[1], 'script')).replace(/<\/script/gi, '<\\/script');
  // Catch accidental corruption during inlining before publishing the file.
  new Function(javascript);
  html = html.replace(match[0], () => `<script type="module">${javascript}</script>`);
}
if (!assets.some(asset => asset.endsWith('.js')) || !assets.some(asset => asset.endsWith('.css'))) {
  throw new Error('Expected JavaScript and CSS assets were not found.');
}
await writeFile(index, html);
for (const asset of assets) await unlink(asset);
if (!(await readdir(resolve(dist, 'assets'))).length) await rmdir(resolve(dist, 'assets'));
console.log(`Inlined ${assets.length} assets into ${basename(index)}`);
