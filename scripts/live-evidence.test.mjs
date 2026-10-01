import fs from 'node:fs';
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fingerprint, verifyEvidence } from './verify-live-evidence.mjs';

// Exercise the verifier with a current-source fixture so PR builds can produce
// candidates before live testing. The publish job checks the unmodified evidence.
const readEvidence = () => {
  const evidence = JSON.parse(fs.readFileSync(new URL('../validation/outlook-classic-live.json', import.meta.url), 'utf8'));
  evidence.sourceFingerprint = fingerprint();
  evidence.baseVersion = JSON.parse(fs.readFileSync(new URL('../plugins/outlook-classic/package/plugin.json', import.meta.url), 'utf8')).version;
  return evidence;
};
test('complete live evidence is accepted when bound to current sources', () => verifyEvidence(readEvidence()));
test('publication rejects stale or incomplete live evidence', () => {
  for (const change of [
    e => { e.sourceFingerprint = '0'.repeat(64); },
    e => { e.baseVersion = '0.0.0'; },
    e => { e.report.passed = false; },
    e => { e.report.syntheticStoresDetached = false; },
    e => { e.report.checks.pop(); },
    e => { e.report.checks[0].status = 'failed'; },
    e => { e.report.metrics.realMessagesRead = 0; },
    e => { e.report.runtimeSha256 = 'invalid'; }
  ]) {
    const evidence = readEvidence();
    change(evidence);
    assert.throws(() => verifyEvidence(evidence));
  }
});
