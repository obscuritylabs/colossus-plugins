import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';

// Used only by the opt-in live test. Never log tool response bodies or handles.
export class LiveClient {
  constructor(executable) {
    this.pending = new Map();
    this.nextId = 1;
    this.stderrBytes = 0;
    this.child = spawn(executable, ['--stdio'], { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
    this.child.stderr.on('data', data => { this.stderrBytes += data.length; });
    this.exited = new Promise(resolve => this.child.on('exit', (code, signal) => {
      this.exit = { code, signal };
      for (const entry of this.pending.values()) entry.reject(new Error('MCP process exited'));
      this.pending.clear();
      resolve(this.exit);
    }));
    this.child.on('error', () => {
      for (const entry of this.pending.values()) entry.reject(new Error('MCP process launch failed'));
    });
    this.lines = createInterface({ input: this.child.stdout });
    this.lines.on('line', line => {
      try {
        const value = JSON.parse(line);
        assert.equal(value.jsonrpc, '2.0');
        const entry = this.pending.get(value.id);
        if (entry) { this.pending.delete(value.id); entry.resolve(value); }
      } catch {
        for (const entry of this.pending.values()) entry.reject(new Error('Invalid MCP framing'));
      }
    });
  }
  request(method, params) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error('MCP request deadline')); }, 30000);
      this.pending.set(id, {
        resolve: value => { clearTimeout(timer); resolve(value); },
        reject: error => { clearTimeout(timer); reject(error); }
      });
      this.child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    });
  }
  async initialize() {
    const response = await this.request('initialize', {
      protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'colossus-live-check', version: '1.0.0' }
    });
    assert.ok(response.result?.capabilities.tools);
    this.child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
  }
  async call(name, args = {}) {
    const response = await this.request('tools/call', { name, arguments: args });
    assert.ok(!response.error && !response.result?.isError, `Tool failed: ${name}`);
    const content = response.result?.content?.find(c => c.type === 'text');
    assert.ok(content, `Missing result: ${name}`);
    return JSON.parse(content.text);
  }
  async rejected(name, args) {
    const response = await this.request('tools/call', { name, arguments: args });
    assert.ok(response.error || response.result?.isError, `Expected rejection: ${name}`);
  }
  async close() {
    this.child.stdin.end();
    const timeout = setTimeout(() => this.child.kill(), 5000);
    await this.exited;
    clearTimeout(timeout);
    this.lines.close();
    assert.equal(this.stderrBytes, 0, 'Server emitted diagnostic data');
    assert.equal(this.exit.code, 0, 'Server did not shut down cleanly');
  }
}
