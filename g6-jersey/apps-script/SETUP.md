# G6 Jersey Configurator: switch on online saving

The page is hosted on GitHub Pages. Entries are saved in a Google Sheet that you own,
through a small Google Apps Script. Setup takes about 5 minutes and is done once.

## 1. Create the Sheet and paste the script

1. Go to https://sheets.new and name the sheet, e.g. **G6 Jersey Orders**.
2. In the sheet, open **Extensions > Apps Script**.
3. Delete what is in `Code.gs`, then paste the full contents of [`Code.gs`](Code.gs) from this folder.
4. Click **Save** (disk icon).

## 2. Set your organiser (admin) key

1. In Apps Script, click the **gear icon (Project Settings)** on the left.
2. Scroll to **Script Properties** > **Add script property**.
3. Property: `ADMIN_KEY` · Value: a password only you know (e.g. a long phrase).
4. Click **Save script properties**.

You type this key on the page (footer > **Organiser login**) to see the size totals,
edit or delete any entry, and download the PDF and CSV.

## 3. Deploy as a web app

1. Click **Deploy > New deployment**.
2. Click the gear next to "Select type" and choose **Web app**.
3. Execute as: **Me** · Who has access: **Anyone**.
4. Click **Deploy**, then **Authorize access** and allow it with your Google account.
   (If Google warns "app isn't verified", click **Advanced > Go to … (unsafe)**. It is your own script.)
5. Copy the **Web app URL** (ends with `/exec`).

## 4. Connect the page

Edit [`../config.js`](../config.js) on GitHub (pencil icon), paste the URL between the quotes:

```js
window.G6_CONFIG = {
  apiUrl: "https://script.google.com/macros/s/XXXXXXXX/exec"
};
```

Commit the change. GitHub Pages updates within a minute or two.

## Notes

- **Several people at once:** every save takes a lock in the script, so simultaneous entries are all kept.
- **Changing sizes:** each person picks a 4–6 digit PIN when they submit. They use **Edit / change size**
  on their card and enter the PIN. The same browser remembers it. PINs are stored only as a hash.
- **Deadline:** 17 Nov 2026, 11:59 PM UAE time, enforced by the script as well as the page.
  Change `DEADLINE` at the top of `Code.gs` if needed, then **Deploy > Manage deployments > Edit > New version**.
- **If you edit Code.gs later**, always redeploy as a *new version* of the same deployment so the URL stays the same.
- The Sheet itself is your raw data. Don't rename the `Orders` tab or its header row.

## Payment proofs (Google Drive)

- Members can attach a screenshot/photo of their payment (form or **Upload payment proof** on their card, using their PIN).
  The entry is then marked **Paid** automatically.
- Files are saved in a private Drive folder named **G6 Jersey Payment Proofs** (created on first upload). Only you can open them.
- On the entries page (Settings), the **Proof → View** button opens the file in Google Drive.
- After pasting the new Code.gs and deploying a new version, Google asks you to allow Drive access. If uploads fail with a
  Drive error, open Apps Script, choose the function **authorizeDrive** in the toolbar, click **Run**, and allow access.
