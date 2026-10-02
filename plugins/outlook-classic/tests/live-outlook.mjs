import assert from 'node:assert/strict';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { LiveClient } from './live-client.mjs';

const [executable, fixturePath, outputPath] = process.argv.slice(2);
assert.ok(executable && fixturePath && outputPath, 'Usage: live-outlook.mjs <exe> <fixture.json> <report.json>');
const fixture = JSON.parse(fs.readFileSync(fixturePath, 'utf8').replace(/^\uFEFF/, ''));
assert.equal(fixture.syntheticOnly, true);
const report = { startedAt: new Date().toISOString(), checks: [], metrics: {}, personalContentLogged: false };
const hash = text => createHash('sha256').update(text).digest('hex');
const handle = (kind, storeId, entryId) => {
  const key = createHash('sha256').update(Buffer.from(storeId, 'hex')).digest('hex').slice(0, 32).toUpperCase();
  const payload = `${kind === 'folder' ? 'f2' : 'm2'}.${key}.${entryId.toUpperCase()}`;
  return `${payload}.${hash(payload).slice(0, 8).toUpperCase()}`;
};
const mailHandle = message => handle('message', fixture.stores[0].storeId, message.entryId);
const folderHandle = (store, folder) => handle('folder', store.storeId, folder.entryId);
const primary = fixture.stores[0];
const mainFolder = folderHandle(primary, primary.folders.mail);
let client = new LiveClient(executable);
await client.initialize();
async function check(name, action) {
  const started = Date.now();
  try {
    await action();
    report.checks.push({ name, status: 'passed', elapsedMs: Date.now() - started });
    console.log(`PASS ${name}`);
  } catch (error) {
    // Never copy assertion expected/actual fields or tool payloads into the report.
    report.checks.push({ name, status: 'failed', errorType: error.name, elapsedMs: Date.now() - started });
    console.log(`FAIL ${name} (${error.name})`);
  }
}
try {
  await check('live COM status uses STA and reports classic Outlook', async () => {
    const status = await client.call('get_status');
    assert.equal(status.connected, true);
    assert.equal(status.apartment, 'STA');
    assert.equal(status.readOnly, false);
    assert.equal(status.supportsSend, false);
    assert.equal(status.supportsPermanentDelete, false);
    report.metrics.outlookVersion = status.version;
    report.metrics.storeCount = status.storeCount;
  });
  await check('store pagination discovers both synthetic PSTs without duplicates', async () => {
    let offset = 0;
    const all = [];
    for (let page = 0; page < 100; page++) {
      const result = await client.call('list_stores', { limit: 3, offset });
      assert.ok(result.items.length <= 3);
      all.push(...result.items.map(item => item.folderHandle));
      if (result.nextOffset === null) break;
      assert.ok(result.nextOffset > offset);
      offset = result.nextOffset;
    }
    assert.equal(new Set(all).size, all.length);
    for (const store of fixture.stores) assert.ok(all.includes(handle('folder', store.storeId, store.rootEntryId)));
    assert.equal(all.length, report.metrics.storeCount);
  });
  await check('folder enumeration and paging preserve Unicode and nested folders', async () => {
    let offset = 0;
    const names = [];
    for (let page = 0; page < 50; page++) {
      const result = await client.call('list_folders', { folderHandle: handle('folder', primary.storeId, primary.rootEntryId), limit: 2, offset });
      assert.ok(result.items.length <= 2);
      names.push(...result.items.map(item => item.name));
      if (result.nextOffset === null) break;
      assert.ok(result.nextOffset > offset);
      offset = result.nextOffset;
    }
    assert.ok(names.includes(primary.folders.unicode.name));
    const nested = await client.call('list_folders', { folderHandle: mainFolder, limit: 50 });
    assert.ok(nested.items.some(item => item.name === 'Nested synthetic folder'));
  });
  await check('store default-folder discovery exposes the synthetic Inbox, Drafts, and Deleted Items', async () => {
    const result = await client.call('get_mail_folders', { storeFolderHandle: handle('folder', primary.storeId, primary.rootEntryId) });
    for (const role of ['inbox', 'drafts', 'deletedItems'])
      assert.ok(result.items.some(item => item.role === role && item.folderHandle));
    await client.rejected('get_mail_folders', { storeFolderHandle: mainFolder });
  });
  await check('message paging returns only mail items with no duplicates', async () => {
    const all = [];
    let offset = 0;
    for (let page = 0; page < 10; page++) {
      const result = await client.call('list_messages', { folderHandle: mainFolder, limit: 1, offset });
      assert.ok(result.items.length <= 1 && result.scanned <= 500);
      all.push(...result.items);
      if (result.nextOffset === null) break;
      assert.ok(result.nextOffset > offset);
      offset = result.nextOffset;
    }
    assert.equal(all.length, fixture.messages.length);
    assert.equal(new Set(all.map(item => item.messageHandle)).size, all.length);
    for (const message of fixture.messages) {
      const result = all.find(item => item.messageHandle === mailHandle(message));
      assert.ok(result);
      assert.equal(result.subject, message.subject);
      assert.equal(result.unread, message.unread);
      assert.ok(Number.isFinite(Date.parse(result.receivedAt)));
    }
  });
  await check('literal Unicode and apostrophe subject filters do not become queries', async () => {
    const match = await client.call('search_messages', { folderHandle: mainFolder, subjectContains: "O'Brien", limit: 50 });
    assert.equal(match.items.length, 1);
    const unicode = await client.call('search_messages', { folderHandle: mainFolder, subjectContains: '\u90ae\u4ef6', limit: 50 });
    assert.equal(unicode.items.length, 1);
    const noMatch = await client.call('search_messages', { folderHandle: mainFolder, subjectContains: "' OR 1=1 --", limit: 50 });
    assert.equal(noMatch.items.length, 0);
  });
  await check('unread-only filter returns the expected unread fixtures', async () => {
    const result = await client.call('search_messages', { folderHandle: mainFolder, unreadOnly: true, limit: 50 });
    assert.equal(result.items.length, fixture.messages.filter(item => item.unread).length);
    assert.ok(result.items.every(item => item.unread === true));
  });
  await check('empty folders and offsets past the end return empty terminal pages', async () => {
    for (const args of [
      { folderHandle: folderHandle(primary, primary.folders.empty) },
      { folderHandle: mainFolder, offset: 1000000 }
    ]) {
      const result = await client.call('search_messages', args);
      assert.equal(result.items.length, 0);
      assert.equal(result.nextOffset, null);
    }
  });
  await check('large-folder scan stops at 500 and resumes on the next page', async () => {
    const args = { folderHandle: folderHandle(primary, primary.folders.bulk), subjectContains: 'NEVER-MATCH-FIXTURE', limit: 50 };
    const first = await client.call('search_messages', args);
    assert.equal(first.items.length, 0);
    assert.equal(first.scanned, 500);
    assert.equal(first.nextOffset, 500);
    const second = await client.call('search_messages', { ...args, offset: first.nextOffset });
    assert.equal(second.scanned, fixture.bulkCount - 500);
    assert.equal(second.nextOffset, null);
    const limited = await client.call('search_messages', { folderHandle: args.folderHandle, limit: 50 });
    assert.equal(limited.items.length, 50);
    report.metrics.bulkMessages = fixture.bulkCount;
  });
  await check('body omission, content bounds, and untrusted-data marker work on live items', async () => {
    for (const message of fixture.messages) {
      const messageHandle = mailHandle(message);
      const omitted = await client.call('get_message', { messageHandle, maxBodyChars: 0 });
      assert.equal(omitted.body, '');
      assert.equal(omitted.bodyIncluded, false);
      const result = await client.call('get_message', { messageHandle, maxBodyChars: 24000 });
      assert.ok(result.body.length <= 24000);
      assert.equal(result.contentTrust, 'untrusted-email');
      assert.equal(result.bodyIncluded, true);
      assert.equal(result.bodyTruncated, message.bodyLength > 24000);
      assert.equal(hash(result.body), message.expectedPrefixSha256);
      assert.equal(result.unread, message.unread);
    }
  });
  await check('Unicode truncation does not split an emoji surrogate pair', async () => {
    const cut = fixture.messages[0].unicodeCut;
    assert.ok(cut > 0);
    const result = await client.call('get_message', { messageHandle: mailHandle(fixture.messages[0]), maxBodyChars: cut });
    assert.equal(result.body.length, cut - 1);
    assert.ok(!result.body.includes('\uFFFD'));
    assert.equal(result.bodyTruncated, true);
  });
  await check('attachment metadata and pagination preserve names without exporting files', async () => {
    const messageHandle = mailHandle(fixture.messages[0]);
    const first = await client.call('list_attachments', { messageHandle, limit: 1 });
    const second = await client.call('list_attachments', { messageHandle, limit: 1, offset: first.nextOffset });
    const names = [...first.items, ...second.items].map(item => item.name).sort();
    assert.deepEqual(names, fixture.attachmentNames.sort());
    assert.equal(second.nextOffset, null);
    const past = await client.call('list_attachments', { messageHandle, offset: 1000000 });
    assert.equal(past.items.length, 0);
    assert.equal(past.nextOffset, null);
    const empty = await client.call('list_attachments', { messageHandle: mailHandle(fixture.messages[1]) });
    assert.equal(empty.items.length, 0);
  });
  await check('non-mail and nonexistent COM item handles fail safely', async () => {
    await client.rejected('get_message', { messageHandle: mailHandle({ entryId: fixture.nonMailEntryId }) });
    await client.rejected('get_message', { messageHandle: mailHandle({ entryId: 'AABBCCDD' }) });
    const valid = mailHandle(fixture.messages[0]);
    assert.ok(valid.length < 160, 'Synthetic handles must remain practical for agent reuse');
    await client.rejected('get_message', { messageHandle: valid.slice(0, -1) + (valid.endsWith('A') ? 'B' : 'A') });
    await client.rejected('get_message', { messageHandle: handle('message', 'AABB', fixture.messages[0].entryId) });
  });
  await check('handles from a second PST resolve within the correct store', async () => {
    const other = fixture.stores[1];
    const result = await client.call('search_messages', { folderHandle: folderHandle(other, other.folders.mail) });
    assert.equal(result.items.length, 1);
    const message = await client.call('get_message', { messageHandle: result.items[0].messageHandle, maxBodyChars: 0 });
    assert.equal(message.subject, 'Secondary PST synthetic message');
  });
  let draft;
  let draftFolder;
  let deletedFolder;
  const archiveFolder = folderHandle(primary, primary.folders.archive);
  const emptyFolder = folderHandle(primary, primary.folders.empty);
  await check('create and update an unsent draft only in the selected synthetic PST', async () => {
    const roles = (await client.call('get_mail_folders', { storeFolderHandle: handle('folder', primary.storeId, primary.rootEntryId) })).items;
    draftFolder = roles.find(item => item.role === 'drafts').folderHandle;
    deletedFolder = roles.find(item => item.role === 'deletedItems').folderHandle;
    draft = await client.call('create_draft', {
      storeFolderHandle: handle('folder', primary.storeId, primary.rootEntryId),
      to: 'test@example.invalid', subject: fixture.mutationSubject, body: 'Synthetic unsent draft.'
    });
    assert.equal(draft.folderHandle, draftFolder);
    assert.equal(draft.sent, false);
    const listed = await client.call('list_messages', { folderHandle: draftFolder, limit: 50 });
    assert.ok(listed.items.some(item => item.messageHandle === draft.messageHandle));
    await client.rejected('update_draft', { messageHandle: draft.messageHandle, subject: 'x'.repeat(513) });
    const updated = await client.call('update_draft', {
      messageHandle: draft.messageHandle, subject: fixture.mutationSubject + ' edited', body: 'Updated synthetic draft.'
    });
    assert.equal(updated.subject, fixture.mutationSubject + ' edited');
    assert.equal(updated.sent, false);
    draft = updated;
    report.metrics.syntheticDraftCreated = 1;
  });
  await check('read-state changes require the current synthetic source folder', async () => {
    await client.rejected('mark_message_read', { messageHandle: draft.messageHandle, sourceFolderHandle: mainFolder, read: true });
    const unread = await client.call('mark_message_read', { messageHandle: draft.messageHandle, sourceFolderHandle: draftFolder, read: false });
    assert.equal(unread.read, false);
    draft.messageHandle = unread.messageHandle;
    const read = await client.call('mark_message_read', { messageHandle: draft.messageHandle, sourceFolderHandle: draftFolder, read: true });
    assert.equal(read.read, true);
    draft.messageHandle = read.messageHandle;
  });
  await check('move and archive stay inside the selected synthetic store', async () => {
    await client.rejected('move_message', {
      messageHandle: draft.messageHandle, sourceFolderHandle: draftFolder,
      destinationFolderHandle: folderHandle(fixture.stores[1], fixture.stores[1].folders.mail)
    });
    const moved = await client.call('move_message', {
      messageHandle: draft.messageHandle, sourceFolderHandle: draftFolder, destinationFolderHandle: emptyFolder
    });
    assert.equal(moved.folderHandle, emptyFolder);
    await client.rejected('update_draft', { messageHandle: moved.messageHandle, subject: 'not a draft' });
    const archived = await client.call('archive_message', {
      messageHandle: moved.messageHandle, sourceFolderHandle: emptyFolder, archiveFolderHandle: archiveFolder
    });
    assert.equal(archived.folderHandle, archiveFolder);
    assert.ok((await client.call('list_messages', { folderHandle: archiveFolder })).items.some(item => item.messageHandle === archived.messageHandle));
    draft.messageHandle = archived.messageHandle;
  });
  await check('delete moves to Deleted Items and refuses permanent deletion', async () => {
    await client.rejected('delete_message', { messageHandle: draft.messageHandle, sourceFolderHandle: mainFolder });
    const deleted = await client.call('delete_message', { messageHandle: draft.messageHandle, sourceFolderHandle: archiveFolder });
    assert.equal(deleted.folderHandle, deletedFolder);
    assert.equal(deleted.permanent, false);
    assert.ok((await client.call('list_messages', { folderHandle: deletedFolder })).items.some(item => item.messageHandle === deleted.messageHandle));
    await client.rejected('delete_message', { messageHandle: deleted.messageHandle, sourceFolderHandle: deletedFolder });
  });
  // Explicitly opted in by the local runner. Content is transient, never output or saved.
  if (fixture.liveInbox) await check('default Inbox bounded live reads preserve unread flags (no content logged)', async () => {
    const liveFolder = handle('folder', fixture.liveInbox.storeId, fixture.liveInbox.entryId);
    const page = await client.call('search_messages', { folderHandle: liveFolder, unreadOnly: true, limit: 2 });
    const additional = await client.call('search_messages', { folderHandle: liveFolder, limit: 1 });
    const messages = [...new Map([...page.items, ...additional.items].map(item => [item.messageHandle, item])).values()];
    assert.ok(messages.length > 0, 'No real mail available');
    for (const item of messages) {
      const read = await client.call('get_message', { messageHandle: item.messageHandle, maxBodyChars: 128 });
      assert.ok(read.body.length <= 128);
      assert.equal(read.unread, item.unread);
      const attachments = await client.call('list_attachments', { messageHandle: item.messageHandle, limit: 2 });
      assert.ok(attachments.items.length <= 2);
      const after = await client.call('get_message', { messageHandle: item.messageHandle, maxBodyChars: 0 });
      assert.equal(after.unread, item.unread);
    }
    report.metrics.realMessagesRead = messages.length;
    report.metrics.realUnreadMessagesChecked = messages.filter(item => item.unread).length;
  });
  await check('MCP shutdown leaves Outlook available and handles survive a new process', async () => {
    await client.close();
    client = new LiveClient(executable);
    await client.initialize();
    assert.equal((await client.call('get_status')).connected, true);
    const read = await client.call('get_message', { messageHandle: mailHandle(fixture.messages[0]), maxBodyChars: 0 });
    assert.equal(read.unread, fixture.messages[0].unread);
  });
  await check('all synthetic unread flags still match after the live tool suite', async () => {
    for (const message of fixture.messages) {
      const read = await client.call('get_message', { messageHandle: mailHandle(message), maxBodyChars: 0 });
      assert.equal(read.unread, message.unread);
    }
  });
} finally {
  await check('MCP exits cleanly with no diagnostic output', () => client.close());
  report.completedAt = new Date().toISOString();
  report.passed = report.checks.every(item => item.status === 'passed');
  fs.writeFileSync(outputPath, JSON.stringify(report, null, 2) + '\n');
  if (!report.passed) process.exitCode = 1;
}
