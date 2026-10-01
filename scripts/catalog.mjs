import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import Ajv2020 from 'ajv/dist/2020.js';
import addFormats from 'ajv-formats';
import semver from 'semver';

const readJson = file => JSON.parse(fs.readFileSync(file, 'utf8').replace(/^\uFEFF/, ''));
const schema = readJson(new URL('../schemas/catalog.schema.json', import.meta.url));
const ajv = new Ajv2020({ allErrors: true });
addFormats(ajv);
const checkSchema = ajv.compile(schema);

export function validateCatalog(catalog) {
  assert.ok(checkSchema(catalog), ajv.errorsText(checkSchema.errors));
  const names = new Set();
  for (const plugin of catalog.plugins) {
    assert.ok(!names.has(plugin.name), 'Duplicate plugin name');
    names.add(plugin.name);
    assert.equal(plugin.sourcePath, `plugins/${plugin.name}`);
    const versions = new Set();
    for (const release of plugin.releases) {
      assert.ok(semver.valid(release.version), 'Invalid release SemVer');
      assert.ok(semver.valid(release.minColossusVersion), 'Invalid Colossus SemVer');
      assert.equal(release.channel, semver.prerelease(release.version) ? 'preview' : 'stable');
      assert.ok(!versions.has(release.version), 'Duplicate release version');
      versions.add(release.version);
      const platforms = new Set();
      for (const artifact of release.artifacts) {
        const platform = `${artifact.platform.os}/${artifact.platform.architecture}`;
        assert.ok(!platforms.has(platform), 'Duplicate artifact platform');
        platforms.add(platform);
        if (plugin.name === 'outlook-classic') assert.equal(platform, 'windows/amd64');
      }
    }
  }
  return catalog;
}

export function addRelease(catalog, manifest, evidence, releasedAt = new Date().toISOString()) {
  validateCatalog(catalog);
  assert.equal(manifest.name, 'outlook-classic');
  assert.equal(manifest.version, evidence.version, 'Manifest version does not match release');
  assert.equal(evidence.checksPassed, true, 'Build checks required');
  assert.equal(evidence.verified, true, 'Registry round trip and trust verification required');
  assert.match(evidence.digest, /^sha256:[a-f0-9]{64}$/);
  assert.equal(evidence.repository, 'ghcr.io/obscuritylabs/colossus-plugin-outlook-classic');
  const result = structuredClone(catalog);
  let plugin = result.plugins.find(p => p.name === manifest.name);
  if (!plugin) {
    plugin = { name: manifest.name, displayName: 'Outlook Classic', description: manifest.description,
      sourcePath: `plugins/${manifest.name}`, releases: [] };
    result.plugins.push(plugin);
  }
  const release = {
    version: manifest.version, channel: semver.prerelease(manifest.version) ? 'preview' : 'stable',
    releasedAt, sourceCommit: evidence.sourceCommit, agentPluginsVersion: '1.0.0',
    minColossusVersion: evidence.minColossusVersion, yanked: false,
    artifacts: [{ reference: `${evidence.repository}@${evidence.digest}`,
      platform: { os: 'windows', architecture: 'amd64', minimumOsVersion: 'Windows 11' },
      requirements: [
        'Classic Outlook for Windows x64, already running with a configured MAPI profile.',
        'Same logged-in interactive Windows user session as Outlook; not a Windows service or container.',
        'Direct COM is unsupported inside the tested Colossus windows_job/AppContainer boundary.',
        'Read-only alpha: mailbox integration coverage is incomplete; no sending, editing, or attachment downloads.',
        'OCI package signed with GitHub OIDC; Windows executable is not Authenticode-signed.'
      ] }]
  };
  const existing = plugin.releases.find(r => r.version === release.version);
  if (existing) {
    assert.deepEqual({ ...existing, releasedAt, yanked: false }, release, 'Published versions are immutable');
    return result; // Idempotent retry; do not un-yank or change timestamps.
  }
  plugin.releases.push(release);
  plugin.releases.sort((a, b) => semver.rcompare(a.version, b.version));
  result.plugins.sort((a, b) => a.name.localeCompare(b.name));
  result.revision++;
  return validateCatalog(result);
}

if (process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href) {
  const [command, releaseDirectory = 'dist/release', catalogPath = 'catalog.json'] = process.argv.slice(2);
  const catalog = readJson(catalogPath);
  if (command === 'validate') {
    validateCatalog(catalog);
    console.log(`Catalog valid: revision ${catalog.revision}, ${catalog.plugins.length} plugins.`);
  } else if (command === 'add') {
    const result = addRelease(catalog, readJson(path.join(releaseDirectory, 'plugin.json')),
      readJson(path.join(releaseDirectory, 'release.json')));
    fs.writeFileSync(catalogPath, JSON.stringify(result, null, 2) + '\n');
  } else throw new Error('Usage: node scripts/catalog.mjs validate | add [release-directory] [catalog-path]');
}
