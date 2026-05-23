# Minding Art CRO Tracker

This project deploys best as a Google Apps Script web app attached to a Google Sheet. The sheet stores the tracker JSON in `CRO_Data!A1` and appends every change to `Activity_Log`.

## Recommended Deployment

1. Create a new Google Sheet named `Minding Art CRO Tracker`.
2. Open `Extensions -> Apps Script`.
3. Replace the default `Code.gs` content with this repo's `Code.gs`.
4. Add a new HTML file named `tracker`.
5. Paste this repo's `tracker.html` into that `tracker` file.
6. Optional: open `Project Settings -> Show appsscript.json manifest file`, then replace the manifest with this repo's `appsscript.json`.
7. Click `Save`.
8. Click `Deploy -> New deployment`.
9. Select `Web app`.
10. Set:
    - `Execute as`: `Me`
    - `Who has access`: `Anyone`
11. Click `Deploy`, approve the permissions, and open the web app URL.

The web app URL is the deployed tracker. No separate static hosting is required.

## Vercel Static Deployment

Vercel can host the frontend only. The Google Sheet backend still has to be deployed through Apps Script first.

1. Deploy `Code.gs` as an Apps Script web app using the steps above.
2. Deploy this repo to Vercel.
3. Open the Vercel URL.
4. Paste the Apps Script web app URL into the setup screen.

This repo includes `index.html` and `vercel.json` so the Vercel root URL (`/`) opens `tracker.html` instead of showing `404: NOT_FOUND`.

## Updating an Existing Deployment

After editing `Code.gs` or `tracker.html`:

1. Paste the new code into the Apps Script project.
2. Click `Deploy -> Manage deployments`.
3. Edit the active web app deployment.
4. Choose `New version`.
5. Click `Deploy`.

## Local/Static Fallback

You can still open `tracker.html` directly in a browser. In that mode, the setup screen asks for a deployed Apps Script URL and saves it in browser localStorage.

The single Apps Script web app deployment is preferred because it avoids browser CORS issues and keeps the UI and backend together.
