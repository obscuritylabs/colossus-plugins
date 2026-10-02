import fs from 'node:fs';
import path from 'node:path';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const evidencePath = path.join(root, 'validation/outlook-classic-live.json');
const readJson = file => JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
export function runtimeFingerprint(directory) {
  const files = [];
  function walk(relative = '') {
    for (const entry of fs.readdirSync(path.join(directory, relative), { withFileTypes: true })) {
      const child = relative ? `${relative}/${entry.name}` : entry.name;
      if (entry.isDirectory()) walk(child);
      else if (entry.isFile()) files.push(child);
      else throw new Error('Unsupported runtime entry');
    }
  }
  walk();
  assert.ok(files.includes('outlook-classic-mcp.exe') && files.includes('outlook-classic-mcp.dll'));
  const hash = createHash('sha256');
  for (const file of files.sort()) {
    const digest = createHash('sha256').update(fs.readFileSync(path.join(directory, file))).digest('hex');
    hash.update(file + '\0' + digest + '\0');
  }
  return hash.digest('hex');
}
export function fingerprint() {
  const files = ['global.json', 'NuGet.Config', 'scripts/Build-OutlookClassic.ps1',
    'scripts/Write-DependencyNotices.ps1', 'scripts/Test-OutlookClassicLive.ps1', 'scripts/verify-live-evidence.mjs'];
  function walk(relative) {
    for (const entry of fs.readdirSync(path.join(root, relative), { withFileTypes: true })) {
      if (['bin', 'obj', 'TestResults'].includes(entry.name)) continue;
      const child = `${relative}/${entry.name}`;
      if (entry.isDirectory()) walk(child);
      else if (entry.isFile()) files.push(child);
      else throw new Error(`Unsupported source entry: ${child}`);
    }
  }
  for (const directory of ['plugins/outlook-classic/src', 'plugins/outlook-classic/package', 'plugins/outlook-classic/tests']) walk(directory);
  const hash = createHash('sha256');
  for (const file of files.sort()) {
    // Git checkouts differ in line endings, not the compiled source text.
    const text = fs.readFileSync(path.join(root, file), 'utf8').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n');
    hash.update(file + '\0' + text + '\0');
  }
  return hash.digest('hex');
}
const requiredChecks = [
  'live COM status uses STA and reports classic Outlook',
  'store pagination discovers both synthetic PSTs without duplicates',
  'folder enumeration and paging preserve Unicode and nested folders',
  'store default-folder discovery exposes the synthetic Inbox, Drafts, and Deleted Items',
  'message paging returns only mail items with no duplicates',
  'literal Unicode and apostrophe subject filters do not become queries',
  'unread-only filter returns the expected unread fixtures',
  'empty folders and offsets past the end return empty terminal pages',
  'large-folder scan stops at 500 and resumes on the next page',
  'body omission, content bounds, and untrusted-data marker work on live items',
  'Unicode truncation does not split an emoji surrogate pair',
  'attachment metadata and pagination preserve names without exporting files',
  'non-mail and nonexistent COM item handles fail safely',
  'handles from a second PST resolve within the correct store',
  'create and update an unsent draft only in the selected synthetic PST',
  'read-state changes require the current synthetic source folder',
  'move and archive stay inside the selected synthetic store',
  'delete moves to Deleted Items and refuses permanent deletion',
  'default Inbox bounded live reads preserve unread flags (no content logged)',
  'MCP shutdown leaves Outlook available and handles survive a new process',
  'all synthetic unread flags still match after the live tool suite',
  'MCP exits cleanly with no diagnostic output',
  'independent COM snapshots show unchanged fixture counts, bodies, unread flags, and modification times',
  'only the test-created synthetic draft was removed after Deleted Items verification'
];
function validateReport(report) {
  assert.equal(report.passed, true, 'Live suite must pass');
  assert.equal(report.syntheticStoresDetached, true, 'Fixture cleanup must pass');
  assert.equal(report.personalContentLogged, false);
  assert.ok(report.metrics.realMessagesRead >= 1 && report.metrics.realMessagesRead <= 3);
  assert.equal(report.metrics.bulkMessages, 502);
  assert.equal(report.metrics.syntheticDraftCreated, 1);
  for (const name of requiredChecks)
    assert.equal(report.checks.filter(check => check.name === name && check.status === 'passed').length, 1, `Missing live check: ${name}`);
  assert.ok(report.checks.every(check => check.status === 'passed'));
}
export function verifyEvidence(evidence) {
  assert.equal(evidence.schemaVersion, 1);
  validateReport(evidence.report);
  assert.equal(evidence.sourceFingerprint, fingerprint(), 'Runtime or test sources changed. Run live Outlook tests and record fresh evidence before publishing.');
  const manifest = readJson(path.join(root, 'plugins/outlook-classic/package/plugin.json'));
  assert.equal(evidence.baseVersion, manifest.version);
  assert.match(evidence.report.executableSha256, /^[a-f0-9]{64}$/);
  assert.match(evidence.report.runtimeSha256, /^[a-f0-9]{64}$/);
}
if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const [command, reportPath] = process.argv.slice(2);
  if (command === 'fingerprint') console.log(fingerprint());
  else if (command === 'runtime-fingerprint') console.log(runtimeFingerprint(reportPath));
  else if (command === 'record') {
    const report = readJson(reportPath);
    const build = readJson(path.join(root, '.local/last-outlook-build.json'));
    validateReport(report);
    assert.equal(build.checksSkipped, false);
    assert.equal(report.executableSha256, build.executableSha256, 'Live-tested binary differs from the local build receipt');
    assert.equal(report.runtimeSha256, build.runtimeSha256, 'Live-tested managed DLLs/runtime differ from the local build receipt');
    assert.equal(build.runtimeSha256, runtimeFingerprint(path.dirname(build.executable)), 'Runtime changed after testing');
    assert.equal(build.sourceFingerprint, fingerprint(), 'Source changed since the tested build');
    const manifest = readJson(path.join(root, 'plugins/outlook-classic/package/plugin.json'));
    const evidence = { schemaVersion: 1, sourceFingerprint: build.sourceFingerprint, baseVersion: manifest.version,
      testedVersion: build.version, scope: 'Classic Outlook in the interactive user session; AppContainer and new Outlook are unsupported.', report };
    verifyEvidence(evidence);
    fs.mkdirSync(path.dirname(evidencePath), { recursive: true });
    fs.writeFileSync(evidencePath, JSON.stringify(evidence, null, 2) + '\n');
    console.log('Recorded passing live evidence for the current source and tested executable.');
  } else if (command === 'check') {
    verifyEvidence(readJson(evidencePath));
    console.log('Live Outlook evidence matches the current runtime, dependency, and test sources.');
  } else throw new Error('Usage: verify-live-evidence.mjs fingerprint | record <report.json> | check');
}
