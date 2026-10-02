import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import { addRelease, validateCatalog } from './catalog.mjs';

const empty = { ...JSON.parse(fs.readFileSync(new URL('../catalog.json', import.meta.url))), revision: 0, plugins: [] };
const manifest = { name: 'outlook-classic', version: '0.1.0-alpha.1', description: 'Read-only classic Outlook', keywords: ['read-only'] };
const evidence = { version: manifest.version, digest: `sha256:${'a'.repeat(64)}`, sourceCommit: 'b'.repeat(40),
  minColossusVersion: '0.11.4', checksPassed: true, verified: true,
  repository: 'ghcr.io/obscuritylabs/colossus-plugin-outlook-classic' };
test('verified release is pinned, classified as preview, and retry is idempotent', () => {
  const first = addRelease(empty, manifest, evidence);
  assert.equal(first.revision, 1);
  assert.equal(first.plugins[0].releases[0].channel, 'preview');
  assert.deepEqual(addRelease(first, manifest, evidence), first);
  assert.equal(empty.plugins.length, 0);
});
test('publication requires verified evidence and matching manifest', () => {
  for (const change of [{ verified: false }, { checksPassed: false }, { version: '9.0.0' }, { digest: 'latest' }])
    assert.throws(() => addRelease(empty, manifest, { ...evidence, ...change }));
});
test('existing version cannot move to another digest', () => {
  const first = addRelease(empty, manifest, evidence);
  assert.throws(() => addRelease(first, manifest, { ...evidence, digest: `sha256:${'c'.repeat(64)}` }));
});
test('new write-capable release updates discovery text without changing old release requirements', () => {
  const first = addRelease(empty, manifest, evidence);
  const nextManifest = { ...manifest, version: '0.1.0-alpha.4',
    description: 'List and manage Outlook mail and unsent drafts', keywords: ['mcp'] };
  const nextEvidence = { ...evidence, version: nextManifest.version, digest: `sha256:${'d'.repeat(64)}` };
  const next = addRelease(first, nextManifest, nextEvidence);
  assert.equal(next.plugins[0].description, nextManifest.description);
  assert.match(next.plugins[0].releases[0].artifacts[0].requirements.join(' '), /unsent drafts/);
  assert.match(next.plugins[0].releases[1].artifacts[0].requirements.join(' '), /Read-only alpha/);
});
test('reject duplicate identities, invalid versions, stable prereleases, and unsupported Outlook platforms', () => {
  const first = addRelease(empty, manifest, evidence);
  for (const mutate of [
    c => c.plugins.push(structuredClone(c.plugins[0])),
    c => c.plugins[0].releases.push(structuredClone(c.plugins[0].releases[0])),
    c => { c.plugins[0].releases[0].version = 'tomorrow'; },
    c => { c.plugins[0].releases[0].channel = 'stable'; },
    c => { c.plugins[0].releases[0].artifacts[0].platform.os = 'linux'; },
    c => { c.plugins[0].releases[0].artifacts[0].reference = `${evidence.repository}:latest`; }
  ]) { const bad = structuredClone(first); mutate(bad); assert.throws(() => validateCatalog(bad)); }
});
