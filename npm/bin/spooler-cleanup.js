#!/usr/bin/env node
'use strict';

// Thin CLI around the bundled SpoolerCleanup.exe. Windows only - the print spooler and the
// spool folder do not exist elsewhere; package.json also declares "os": ["win32"].

const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const pkg = require('../package.json');
const exe = path.join(__dirname, 'SpoolerCleanup.exe');

function die(message) {
  console.error('spooler-cleanup: ' + message);
  process.exit(1);
}

if (process.platform !== 'win32') {
  die('Windows only - the print spooler and the spool folder do not exist on ' + process.platform + '.');
}

const argv = process.argv.slice(2);

// the executable has no /version flag at all, and spells help as /help
if (argv[0] === '--version' || argv[0] === '-v') {
  console.log(pkg.version);
  process.exit(0);
}
const args = argv.map((a) => (a === '--help' || a === '-h') ? '/help' : a);

if (!fs.existsSync(exe)) {
  die('the bundled executable is missing (' + exe + '). Reinstall the package.');
}

// stdio: inherit - the app writes to whatever handle it was given (console or pipe), so
// `spooler-cleanup /list | findstr` works; the exit code is passed through unchanged
// (0 = ok, 2 = needs elevation, 1 = failure).
const res = spawnSync(exe, args, { stdio: 'inherit' });
if (res.error) die(res.error.message);
process.exit(res.status === null ? 1 : res.status);
