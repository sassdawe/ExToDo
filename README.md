# ExToDo

Export all or selected Microsoft To Do lists and their tasks to `extodo.json`. ExToDo is a browser-only exporter served by a small static Node.js server. It has no database, upload endpoint, server-side Graph calls, or client secret.

## Register the app with Microsoft

1. In the [Microsoft Entra admin center](https://entra.microsoft.com/), create an **App registration**. For supported account types, choose **Accounts in any organizational directory and personal Microsoft accounts**.
2. Under **Authentication**, add a **Single-page application (SPA)** platform. Register `https://<app-host>/auth.html` for the production hostname, and `http://localhost:8080/auth.html` if you also want to test locally. Use exact URLs, including any custom domain users will visit. Do not enable implicit-grant tokens or create a secret. Do not select the **Web** platform for this callback.
3. Under **API permissions**, add **Microsoft Graph → Delegated permissions → Tasks.Read**. Remove the default `User.Read` permission if present; ExToDo does not need it. The app asks for consent when needed. Some work or school tenants require administrator approval under their consent policies.
4. Copy the **Application (client) ID**. This is a public identifier, not a credential.

## Run locally

With Node.js 24:

```sh
npm ci
MICROSOFT_CLIENT_ID='<application-client-id>' npm start
npm test
```

Open `http://localhost:8080`. Allow Microsoft sign-in popups if your browser blocks them.

## Deploy to Azure App Service

1. Build the image with `docker build -t extodo .`, tag it for your container registry, and push it there.
2. Create an Azure App Service **Web App** using a **Linux** custom container. Configure the image from your registry and set the container's target port to **8080**. For an App Service using the classic (non-sidecar) container configuration, set the `WEBSITES_PORT` app setting to `8080` instead.
3. Set `MICROSOFT_CLIENT_ID` in the App Service application settings and enable **HTTPS Only**. Keep the web app publicly reachable; do not enable App Service built-in authentication because Microsoft sign-in is handled by the browser. Confirm that the deployed hostname plus `/auth.html` is registered under the SPA platform.
4. Optionally configure App Service Health check to use `/healthz`.

## Export and privacy

Sign in, consent to read-only access, then check lists and choose **Download selected lists**, or choose **Download all lists**. The output has a top-level `lists` array. Each list contains its Graph list fields and a `tasks` array of the task records returned by Graph, including all pages. The app does not fetch separate related resources such as attachments.

- The browser calls Microsoft Graph directly. The container serves only static assets and the public client ID; it never receives access tokens, task contents, encryption keys, or the finished export. There is no app analytics or task logging.
- While assembling the export, completed chunks are encrypted with AES-256-GCM in browser memory using a fresh, non-extractable key per export and a random IV per chunk. List names and IDs used by the picker and individual Graph responses necessarily exist briefly as plaintext in the browser. Tokens are held in memory, and neither tokens nor task content are written to browser storage by the app.
- The browser decrypts the assembled export to create an ordinary **unencrypted JSON file** for the user. Protect the file, including copies your browser, OS, or cloud sync may retain. JavaScript cannot guarantee erasure of every plaintext copy from browser memory.
- This design keeps task data out of the running service and its logs. It does not protect against a browser extension or an operator who changes the served JavaScript: users still have to trust the code delivered to their browser.
