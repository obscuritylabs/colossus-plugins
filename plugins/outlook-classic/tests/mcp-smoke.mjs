import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';

const executable = process.argv[2];
assert.ok(executable, 'Usage: node mcp-smoke.mjs <path-to-published-exe>');
const child = spawn(executable, ['--stdio'], { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
let nextId = 1;
const pending = new Map();
let stderr = '';
child.stderr.setEncoding('utf8');
child.stderr.on('data', chunk => { stderr = (stderr + chunk).slice(-4000); });
const exited = new Promise(resolve => child.on('exit', code => {
  for (const request of pending.values()) request.reject(new Error(`Server exited ${code}: ${stderr}`));
  resolve(code);
}));
child.on('error', error => { for (const request of pending.values()) request.reject(error); });
const lines = createInterface({ input: child.stdout });
lines.on('line', line => {
  try {
    const message = JSON.parse(line);
    assert.equal(message.jsonrpc, '2.0');
    if (message.id !== undefined && pending.has(message.id)) {
      pending.get(message.id).resolve(message);
      pending.delete(message.id);
    }
  } catch (error) {
    for (const request of pending.values()) request.reject(error);
  }
});

function request(method, params) {
  const id = nextId++;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timeout: ${method}`)); }, 10000);
    pending.set(id, {
      resolve: result => { clearTimeout(timer); resolve(result); },
      reject: error => { clearTimeout(timer); reject(error); }
    });
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
  });
}

try {
  const init = await request('initialize', {
    protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'colossus-plugin-smoke', version: '1.0.0' }
  });
  assert.ok(init.result?.capabilities.tools);
  child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
  assert.deepEqual((await request('ping', {})).result, {});
  const listed = await request('tools/list', {});
  const expected = ['get_status', 'list_stores', 'list_folders', 'get_mail_folders', 'list_messages',
    'search_messages', 'get_message', 'list_attachments', 'mark_message_read', 'move_message',
    'archive_message', 'delete_message', 'create_draft', 'update_draft'];
  assert.deepEqual(listed.result.tools.map(tool => tool.name).sort(), expected.sort());
  const readOnly = new Set(['get_status', 'list_stores', 'list_folders', 'get_mail_folders',
    'list_messages', 'search_messages', 'get_message', 'list_attachments']);
  const destructive = new Set(['move_message', 'archive_message', 'delete_message']);
  for (const tool of listed.result.tools) {
    assert.equal(tool.annotations.readOnlyHint, readOnly.has(tool.name), tool.name);
    assert.equal(tool.annotations.destructiveHint, destructive.has(tool.name), tool.name);
  }
  const invalidLimit = await request('tools/call', { name: 'list_stores', arguments: { limit: 51 } });
  assert.equal(invalidLimit.result?.isError, true);
  assert.match(JSON.stringify(invalidLimit.result), /limit must be between 1 and 50/);
  const invalidHandle = await request('tools/call', { name: 'get_message', arguments: { messageHandle: '邮件-INVALID' } });
  assert.equal(invalidHandle.result?.isError, true);
  assert.match(JSON.stringify(invalidHandle.result), /Invalid Outlook handle/);
  const unknown = await request('tools/call', { name: 'send_email', arguments: {} });
  assert.ok(unknown.error || unknown.result?.isError, 'send tool must not exist');
  for (const name of ['move_message', 'archive_message', 'delete_message', 'mark_message_read', 'update_draft']) {
    const absent = await request('tools/call', { name, arguments: {} });
    assert.ok(absent.error || absent.result?.isError, `${name} must require exact handles`);
  }
  const missing = await request('tools/call', { name: 'get_message', arguments: {} });
  assert.ok(missing.error || missing.result?.isError, 'required handle must be enforced');
  const largeHandle = await request('tools/call', { name: 'get_message', arguments: { messageHandle: 'A'.repeat(24001) } });
  assert.equal(largeHandle.result?.isError, true);
  assert.equal(stderr, '', 'mail/protocol errors must not leak into diagnostic logs');
  console.log('PASS MCP handshake, ping, 14 annotated tools, Unicode/malformed input, bounds, required handles, and unavailable send. No mailbox was opened.');
} finally {
  child.stdin.end();
  const killTimer = setTimeout(() => child.kill(), 3000);
  await exited;
  clearTimeout(killTimer);
  lines.close();
}
