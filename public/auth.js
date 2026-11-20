// The same-origin MSAL redirect bridge returns the OAuth response to the opener.
if (location.hash) void globalThis.msalRedirectBridge.broadcastResponseToMainFrame().catch(() => {});
