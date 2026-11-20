import { exportLists, getLists } from './export.js';

const ui = Object.fromEntries(
  ['connect', 'disconnect', 'status', 'picker', 'lists', 'selected', 'all']
    .map((id) => [id, document.getElementById(id)]),
);
const scopes = ['https://graph.microsoft.com/Tasks.Read'];
let client;
let account;
let busy = false;
let reconnect = false;
let listsLoaded = false;

function checked() {
  return Array.from(ui.lists.querySelectorAll('input:checked'), (input) => input.value);
}

function render() {
  ui.connect.hidden = Boolean(account) && !reconnect && listsLoaded;
  ui.connect.textContent = reconnect ? 'Reconnect to Microsoft' : account ? 'Retry loading lists' : 'Sign in with Microsoft';
  ui.connect.disabled = busy || !client;
  ui.disconnect.hidden = !account;
  ui.disconnect.disabled = busy;
  ui.picker.hidden = !listsLoaded;
  ui.picker.disabled = busy || reconnect;
  ui.selected.disabled = checked().length === 0;
}

function show(message) {
  ui.status.textContent = message;
}

function showError(error, fallback) {
  if (account && (error instanceof globalThis.msal.InteractionRequiredAuthError || error.status === 401 || error.status === 403)) {
    reconnect = true;
    show('Microsoft needs you to reconnect. Your organization may require an administrator to approve access.');
  } else if (error.status === 429) {
    show('Microsoft Graph is limiting requests. Wait a moment and try again.');
  } else if (error.code === 'lists-changed') {
    listsLoaded = false;
    ui.lists.replaceChildren();
    show('A selected list is no longer available. Reload the lists and try again.');
  } else {
    show(fallback);
  }
}

async function loadLists(accessToken) {
  const lists = await getLists(accessToken);
  ui.lists.replaceChildren(...lists.map(({ id, displayName }) => {
    const item = document.createElement('li');
    const label = document.createElement('label');
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.value = id;
    label.append(box, displayName || 'Unnamed list');
    item.append(label);
    return item;
  }));
  listsLoaded = true;
  show(lists.length ? 'Choose which lists to download.' : 'No task lists found. You can download an empty export.');
}

async function run(action, fallback) {
  busy = true;
  render();
  try {
    await action();
  } catch (error) {
    showError(error, fallback);
  } finally {
    busy = false;
    render();
  }
}

ui.connect.addEventListener('click', () => run(async () => {
  show(account ? 'Connecting to Microsoft Graph…' : 'Complete sign-in in the Microsoft window.');
  const response = !account
    ? await client.loginPopup({ scopes, prompt: 'select_account' })
    : reconnect
      ? await client.acquireTokenPopup({ scopes, account, prompt: 'consent' })
      : await client.acquireTokenSilent({ scopes, account });
  account = response.account;
  reconnect = false;
  show('Loading task lists…');
  await loadLists(response.accessToken);
}, 'Could not sign in or load task lists. Make sure popups are allowed and try again.'));

ui.disconnect.addEventListener('click', () => run(async () => {
  const previous = account;
  account = null;
  reconnect = false;
  listsLoaded = false;
  ui.lists.replaceChildren();
  render();
  show('Signing out…');
  try {
    await client.logoutPopup({ account: previous });
    show('Signed out of ExToDo.');
  } finally {
    await client.clearCache();
  }
}, 'Signed out of ExToDo. Microsoft sign-out may not have completed.'));

function download(selectedIds) {
  return run(async () => {
    show('Preparing your download…');
    const { accessToken } = await client.acquireTokenSilent({ scopes, account });
    const file = await exportLists(accessToken, selectedIds);
    const url = URL.createObjectURL(file);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'extodo.json';
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    show('Download requested. Keep the unencrypted JSON file private.');
  }, 'Export failed. No file was downloaded. Please try again.');
}

ui.lists.addEventListener('change', render);
ui.selected.addEventListener('click', () => download(new Set(checked())));
ui.all.addEventListener('click', () => download(null));

try {
  if (!globalThis.crypto?.subtle) throw new Error('Secure context required.');
  const response = await fetch('/config.json');
  if (!response.ok) throw new Error('Configuration unavailable.');
  const { clientId } = await response.json();
  const bridge = new URL('/auth.html', location.origin).href;
  const auth = new globalThis.msal.PublicClientApplication({
    auth: {
      clientId,
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: bridge,
      postLogoutRedirectUri: bridge,
    },
    cache: { cacheLocation: 'memoryStorage' },
  });
  await auth.initialize();
  client = auth;
  show('Sign in to load your task lists.');
} catch {
  show('ExToDo could not start. Use HTTPS (or localhost) and check the app configuration.');
}
render();
