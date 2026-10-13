# Attendance email sender setup

The Gmail button sends the generated Excel workbook to `nepalt12345@gmail.com` through this Apps Script web app. The browser page does not contain mail credentials.

## Configure and deploy

1. Create a project at [script.google.com](https://script.google.com/) and paste `Code.gs` into its editor.
2. Keep `RECIPIENT_EMAIL` set to the intended destination in `Code.gs`.
3. Choose **Deploy > New deployment > Web app**.
4. Set **Execute as** to **User accessing the web app**. Set access to **Anyone with a Google account**. This requires an authorized Google account and keeps sends tied to the account pressing the button.
5. Deploy and approve the requested mail permission. Copy the web app URL ending in `/exec`.
6. In `js/attendance.js`, replace `PASTE_DEPLOYED_APPS_SCRIPT_WEB_APP_URL_HERE` in `GMAIL_SCRIPT_URL` with the copied URL.
7. Publish the updated attendance page. On each device, sign in and authorize the Apps Script with the Google account that should send the reports. The authenticated Google account is the sender.

When the Gmail button is pressed, a new tab displays the send result. The workbook is emailed as an `.xlsx` attachment, with the shift and date in the subject. The sender is limited to 10 reports per day by this script; Google account mail quotas also apply.

Google sign-in and Apps Script authorization authenticate the sender; no sender address is stored in the attendance page.

For code changes to the Apps Script after deployment, create a new deployment version or update the existing deployment to use the new version.
