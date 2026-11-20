const graph = 'https://graph.microsoft.com';
const listsUrl = `${graph}/v1.0/me/todo/lists`;

async function* pages(initialUrl, accessToken, fetchPage) {
  const seen = new Set();
  let next = initialUrl;
  while (next !== undefined) {
    const url = new URL(next);
    if (url.origin !== graph || url.username || url.password || url.hash || seen.has(url.href)) {
      throw new Error('Invalid Graph pagination link.');
    }
    seen.add(url.href);
    const response = await fetchPage(url, {
      headers: { Authorization: ['Bearer', accessToken].join(' ') },
      cache: 'no-store',
      credentials: 'omit',
      redirect: 'error',
    });
    if (!response.ok) {
      throw Object.assign(new Error('Graph request failed.'), { status: response.status });
    }
    const body = await response.json();
    if (!Array.isArray(body.value)) throw new Error('Invalid Graph collection.');
    next = body['@odata.nextLink'];
    yield body.value;
  }
}

export async function getLists(accessToken, fetchPage = fetch) {
  const lists = [];
  for await (const page of pages(listsUrl, accessToken, fetchPage)) {
    for (const { id, displayName } of page) lists.push({ id, displayName });
  }
  return lists;
}

export async function createExportBuffer() {
  let key = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  const chunks = [];
  const encoder = new TextEncoder();
  return {
    async append(text) {
      const iv = crypto.getRandomValues(new Uint8Array(12));
      const bytes = encoder.encode(text);
      try {
        const ciphertext = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes);
        const chunk = { iv, ciphertext };
        chunks.push(chunk);
        return chunk;
      } finally {
        bytes.fill(0);
      }
    },
    async toBlob() {
      const parts = [];
      try {
        for (const { iv, ciphertext } of chunks) {
          parts.push(new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, ciphertext)));
        }
        return new Blob(parts, { type: 'application/json' });
      } finally {
        for (const part of parts) part.fill(0);
        chunks.length = 0;
        key = null;
      }
    },
  };
}

// Keep completed JSON chunks encrypted while fetching the rest of the export.
// The ordinary JSON Blob is created only after every requested Graph page succeeds.
export async function exportLists(accessToken, selectedIds, fetchPage = fetch) {
  const output = await createExportBuffer();
  const remaining = selectedIds === null ? null : new Set(selectedIds);
  await output.append('{"lists":[');
  let listSeparator = '';
  for await (const page of pages(listsUrl, accessToken, fetchPage)) {
    for (const list of page) {
      if (remaining && !remaining.delete(list.id)) continue;
      if (typeof list.id !== 'string' || !list.id || list.id === '.' || list.id === '..') {
        throw new Error('Invalid Graph list ID.');
      }
      await output.append(listSeparator + JSON.stringify({ ...list, tasks: [] }).slice(0, -2));
      let taskSeparator = '';
      const tasksUrl = `${listsUrl}/${encodeURIComponent(list.id)}/tasks`;
      for await (const tasks of pages(tasksUrl, accessToken, fetchPage)) {
        if (tasks.length === 0) continue;
        await output.append(taskSeparator + JSON.stringify(tasks).slice(1, -1));
        taskSeparator = ',';
      }
      await output.append(']}');
      listSeparator = ',';
    }
    if (remaining?.size === 0) break;
  }
  if (remaining?.size) {
    throw Object.assign(new Error('A selected list was removed.'), { code: 'lists-changed' });
  }
  await output.append(']}');
  return output.toBlob();
}
