import assert from 'node:assert/strict';
import test from 'node:test';
import { createExportBuffer, exportLists, getLists } from '../public/export.js';

const root = 'https://graph.microsoft.com/v1.0/me/todo/lists';
const first = { id: 'list+1', displayName: 'Personal 📝', isOwner: true };
const second = { id: 'list2', displayName: 'Empty' };
const task = { id: 'task1', title: 'Private task', body: { content: '<example>' } };
const task2 = { id: 'task2', title: 'Next page' };

function graphMock(pages) {
  const calls = [];
  return {
    calls,
    fetchPage: async (url, options) => {
      calls.push(url.href);
      assert.equal(options.headers.Authorization, ['Bearer', 'mock-access-token'].join(' '));
      assert.equal(options.cache, 'no-store');
      assert.equal(options.redirect, 'error');
      assert.ok(pages.has(url.href), `Unexpected Graph request: ${url.href}`);
      return Response.json(pages.get(url.href));
    },
  };
}

function fixture() {
  return graphMock(new Map([
    [root, { value: [first], '@odata.nextLink': `${root}?$skiptoken=opaque%2B1` }],
    [`${root}?$skiptoken=opaque%2B1`, { value: [second] }],
    [`${root}/list%2B1/tasks`, { value: [task], '@odata.nextLink': `${root}/list%2B1/tasks?$skiptoken=opaque%2B2` }],
    [`${root}/list%2B1/tasks?$skiptoken=opaque%2B2`, { value: [task2] }],
    [`${root}/list2/tasks`, { value: [] }],
  ]));
}

test('lists all Graph pages and keeps only IDs and display names for the picker', async () => {
  assert.deepEqual(await getLists('mock-access-token', fixture().fetchPage), [
    { id: first.id, displayName: first.displayName }, second,
  ]);
});

test('exports all list and task pages, including empty lists', async () => {
  const file = await exportLists('mock-access-token', null, fixture().fetchPage);
  assert.equal(file.type, 'application/json');
  assert.deepEqual(JSON.parse(await file.text()), {
    lists: [{ ...first, tasks: [task, task2] }, { ...second, tasks: [] }],
  });
});

test('exports only selected lists without fetching tasks for other lists', async () => {
  const { calls, fetchPage } = fixture();
  const file = await exportLists('mock-access-token', new Set([second.id]), fetchPage);
  assert.deepEqual(JSON.parse(await file.text()), { lists: [{ ...second, tasks: [] }] });
  assert.ok(calls.every((url) => !url.includes('list%2B1/tasks')));
  await assert.rejects(exportLists('mock-access-token', new Set(['removed']), fetchPage), { code: 'lists-changed' });
});

test('rejects pagination outside Graph or a pagination loop', async () => {
  for (const next of ['https://example.test/collect', `https://user@graph.microsoft.com/v1.0/me`, root]) {
    const { calls, fetchPage } = graphMock(new Map([[root, { value: [], '@odata.nextLink': next }]]));
    await assert.rejects(getLists('mock-access-token', fetchPage));
    assert.deepEqual(calls, [root]);
  }
});

test('does not return a partial download when Graph rejects a task request', async () => {
  const { fetchPage } = fixture();
  await assert.rejects(exportLists('mock-access-token', null, (url, options) => (
    url.pathname.endsWith('/list2/tasks') ? new Response(null, { status: 429 }) : fetchPage(url, options)
  )), { status: 429 });
});

test('buffers encrypted bytes with fresh IVs and produces plaintext only for the download', async () => {
  const buffer = await createExportBuffer();
  const firstChunk = await buffer.append('"private"');
  const secondChunk = await buffer.append('"private"');
  assert.equal(firstChunk.iv.byteLength, 12);
  assert.equal(firstChunk.ciphertext.byteLength, '"private"'.length + 16);
  assert.notDeepEqual(firstChunk.iv, secondChunk.iv);
  assert.notDeepEqual(firstChunk.ciphertext, secondChunk.ciphertext);
  assert.notEqual(new TextDecoder().decode(firstChunk.ciphertext), '"private"');
  assert.equal(await (await buffer.toBlob()).text(), '"private""private"');
  await assert.rejects(buffer.append('later'));
});
