# Render deploy

This deploy path is for a disposable demo where the Web document, editor auth sessions, and Word edit sessions may disappear. Use a Render Free Web Service without a persistent disk.

## Requirements

- A Render account.
- A Microsoft app registration with a client ID and client secret.
- Testers with a Microsoft account, OneDrive access, and a Word editing environment.
- This repository connected to Render from GitHub.

## Render service

1. Create a new Render **Web Service** from this repository.
2. Choose **Docker** as the environment so the included `Dockerfile` installs Node.js dependencies and the `pandoc` CLI.
3. Select the **Free** instance type.
4. Do not attach a persistent disk.
5. Deploy the service.

The Docker image uses Node 22, enables `pnpm@10.24.0`, installs `pandoc`, runs `pnpm build`, and starts Next.js with Render's `PORT`.

## Environment variables

Set these on the Render service:

| Name | Value |
| --- | --- |
| `APP_BASE_URL` | `https://<render-service-name>.onrender.com` |
| `MICROSOFT_CLIENT_ID` | Microsoft app registration application/client ID |
| `MICROSOFT_CLIENT_SECRET` | Microsoft app registration client secret value |
| `WEB_DOCUMENT_DB_PATH` | `/tmp/word-sync-demo/web-document.sqlite` |
| `EDITOR_AUTH_DB_PATH` | `/tmp/word-sync-demo/editor-auth.sqlite` |
| `WORD_EDIT_SESSION_DB_PATH` | `/tmp/word-sync-demo/word-edit-sessions.sqlite` |

Do not commit Microsoft secrets or local SQLite files.

## Microsoft app registration

Configure the Microsoft app registration used by `MICROSOFT_CLIENT_ID`:

1. Add this Web redirect URI:

   ```text
   https://<render-service-name>.onrender.com/auth/callback
   ```

2. Ensure the app registration supports the intended testers. For external demo testers, the registration must allow accounts outside the owning tenant if tenant policy permits it.
3. Add delegated Microsoft Graph permission `Files.ReadWrite.AppFolder`.
4. Create a client secret and copy only its value to `MICROSOFT_CLIENT_SECRET` in Render.

`APP_BASE_URL` and the Microsoft redirect URI must use the same public Render origin.

## Demo limitations on Render Free

- Render Free instances spin down after idle time, and the local filesystem is ephemeral.
- SQLite files under `/tmp/word-sync-demo` can disappear after restarts, redeploys, or spin-downs.
- Existing browser cookies may point to deleted editor sessions after a restart; sign out/in again if that happens.
- Keep the browser tab open while editing in Word. The Web document editor polls every 3 seconds, which helps keep the Free service awake during a normal demo.
- If the service spins down during a Word edit session, finishing the session can fail because the local session record may be gone.

## Validation

Before deploying:

```sh
pnpm lint
pnpm test
pnpm build
docker build -t word-sync-demo .
docker run --rm word-sync-demo pandoc --version
```

After deploying:

1. Open the Render URL and wait for any cold start.
2. Sign in with Microsoft.
3. Start a Word edit session.
4. Open the OneDrive working copy via the Word launch path or OneDrive fallback.
5. Edit and save the working copy.
6. Return to the web page and run the finish/import action.
7. Confirm the Web document version updates.
8. Check Render logs for missing environment variables, missing `pandoc`, Graph permission failures, or redirect URI mismatches.
