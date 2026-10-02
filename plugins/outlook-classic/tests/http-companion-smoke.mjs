import assert from 'node:assert/strict';
import { randomBytes } from 'node:crypto';
import { spawn } from 'node:child_process';
import { request as httpRequest } from 'node:http';
import { createInterface } from 'node:readline';

const executable = process.argv[2];
assert.ok(executable, 'Usage: node http-companion-smoke.mjs <path-to-published-exe>');
const token = randomBytes(32).toString('base64url');
const child = spawn(executable, ['--http-companion'], {
  windowsHide: true,
  stdio: ['pipe', 'pipe', 'pipe'],
});
let stderr = '';
child.stderr.setEncoding('utf8');
child.stderr.on('data', chunk => { stderr = (stderr + chunk).slice(-4000); });
const lines = createInterface({ input: child.stdout });
const ready = new Promise((resolve, reject) => {
  const timeout = setTimeout(() => reject(new Error(`Companion startup timed out: ${stderr}`)), 15000);
  lines.once('line', line => { clearTimeout(timeout); resolve(JSON.parse(line)); });
  child.once('error', error => { clearTimeout(timeout); reject(error); });
  child.once('exit', code => { clearTimeout(timeout); reject(new Error(`Companion exited ${code}: ${stderr}`)); });
});

function post(endpoint, body, headers = {}) {
  return new Promise((resolve, reject) => {
    const request = httpRequest(endpoint, {
      method: 'POST',
      timeout: 10000,
      headers: {
        'Content-Type': 'application/json',
        Accept: 'application/json, text/event-stream',
        'MCP-Protocol-Version': '2025-11-25',
        ...headers,
      },
    }, response => {
      let text = '';
      response.setEncoding('utf8');
      response.on('data', chunk => {
        text += chunk;
        if (text.length > 1024 * 1024) request.destroy(new Error('Oversized MCP response'));
      });
      response.on('end', () => resolve({ status: response.statusCode, body: text }));
    });
    request.on('timeout', () => request.destroy(new Error('MCP request timed out')));
    request.on('error', reject);
    request.end(JSON.stringify(body));
  });
}

const initialize = {
  jsonrpc: '2.0', id: 1, method: 'initialize', params: {
    protocolVersion: '2025-11-25', capabilities: {},
    clientInfo: { name: 'companion-smoke', version: '1' },
  },
};

let completed = false;
try {
  child.stdin.write(`${token}\n`);
  const { endpoint } = await ready;
  const parsed = new URL(endpoint);
  assert.equal(parsed.hostname, '127.0.0.1');
  assert.equal(parsed.pathname, '/mcp');
  assert.equal((await post(endpoint, initialize)).status, 401);
  assert.equal((await post(endpoint, initialize, { Authorization: 'Bearer wrong' })).status, 401);
  const authorization = { Authorization: `Bearer ${token}` };
  assert.equal((await post(endpoint, initialize, { ...authorization, Host: 'evil.example' })).status, 403);
  assert.equal((await post(endpoint, initialize, { ...authorization, Origin: 'https://evil.example' })).status, 403);
  const accepted = await post(endpoint, initialize, authorization);
  assert.equal(accepted.status, 200);
  assert.match(accepted.body, /"tools"/);
  const listed = await post(endpoint, { jsonrpc: '2.0', id: 2, method: 'tools/list', params: {} }, authorization);
  assert.equal(listed.status, 200);
  assert.match(listed.body, /"get_status"/);
  assert.match(listed.body, /"list_messages"/);
  assert.equal(stderr, '');
  completed = true;
  console.log('PASS companion loopback, bearer/Host/Origin denial, MCP discovery. No mailbox was opened.');
} finally {
  child.stdin.end();
  const stopped = child.exitCode === null
    ? new Promise(resolve => child.once('exit', resolve))
    : Promise.resolve(child.exitCode);
  const killTimer = setTimeout(() => child.kill(), 5000);
  const code = await stopped;
  clearTimeout(killTimer);
  lines.close();
  if (completed) assert.equal(code, 0, `Companion shutdown failed: ${stderr}`);
}
