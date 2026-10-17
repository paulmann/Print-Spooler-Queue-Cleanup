#!/usr/bin/env node
'use strict';

// Puts the published release binary into bin/ so that `npm pack` / `npm publish` ship
// exactly the artifact GitHub released - no local build, no drift between the channels.

const fs = require('node:fs');
const path = require('node:path');
const pkg = require('./package.json');

const asset = 'SpoolerCleanup-' + pkg.version + '-x64.exe';
const url = 'https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/download/v' +
            pkg.version + '/' + asset;
const dest = path.join(__dirname, 'bin', 'SpoolerCleanup.exe');

if (fs.existsSync(dest)) {
  console.log('prepare: ' + path.relative(__dirname, dest) + ' already present (' +
              fs.statSync(dest).size + ' bytes)');
  process.exit(0);
}

console.log('prepare: downloading ' + url);
fetch(url, { redirect: 'follow' })
  .then((res) => {
    if (!res.ok) throw new Error('HTTP ' + res.status + ' for ' + url);
    return res.arrayBuffer();
  })
  .then((buf) => {
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.writeFileSync(dest, Buffer.from(buf));
    console.log('prepare: wrote ' + path.relative(__dirname, dest) + ' (' +
                fs.statSync(dest).size + ' bytes)');
  })
  .catch((err) => {
    console.error('prepare: ' + err.message);
    console.error('prepare: publish the GitHub release for v' + pkg.version +
                  ' first, or place the .exe into npm/bin manually.');
    process.exit(1);
  });
