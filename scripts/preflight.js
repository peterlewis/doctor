#!/usr/bin/env node
/** Preflight checks for Doctor (Node >=18 & required deps) */
const fs = require('fs');
const path = require('path');
const required = [
  'kleur',
  'listr',
  'markdown-it',
  '@pnp/cli-microsoft365',
  'fast-glob',
  'gray-matter'
];
let ok = true;
const nodeMajor = Number(process.versions.node.split('.')[0]);
if (nodeMajor < 18) {
  console.error(`❌ Node ${process.versions.node} < 18 (required >=18)`);
  ok = false;
} else {
  console.log(`✅ Node ${process.versions.node}`);
}
for (const pkg of required) {
  try { require.resolve(pkg); console.log(`✅ ${pkg}`); } catch { console.error(`❌ Missing ${pkg}`); ok = false; }
}
if (!fs.existsSync(path.join(process.cwd(), 'doctor.json'))) {
  console.warn('⚠️ doctor.json missing (run `doctor init`).');
}
if (!ok) { console.error('\nPreflight failed. Run `npm install` or upgrade Node.'); process.exit(1); }
console.log('\nPreflight passed.');