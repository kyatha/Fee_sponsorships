# PhD Fee Bridge (demo)

Sponsors cover UMass Amherst graduate application fees ($90) for applicants from African countries. Applicants post a short research statement. Sponsors either cover a whole fee or chip in part and pool with others. Payments in this demo are **simulated**: nothing is charged and no card details are stored.

There are no logins. Applicants come back with their **application number + email**, sponsors with their **pledge code + email**, and committed sponsors with their **sponsor code + email**.

Not affiliated with or endorsed by UMass Amherst.

## How it works

- **Whole fee or pool:** a sponsor can cover the full $90, or chip in any amount from $5 up, for example $10, $30 or $50. The board's "Pools waiting to fill" list shows the pools closest to full first.
- **Pledge first, pay when full:** nobody pays until a fee is fully pledged. Then every sponsor on that fee has 24 hours to pay their share. If someone doesn't pay, only their part opens up again.
- **Pools that stall:** if a pool hasn't filled within 7 days, or the deadline is less than 2 days away, its unpaid pledges are released and those sponsors are emailed.
- **One open pledge at a time,** unless you're a committed sponsor, who can hold as many as the number of students they committed to.
- **Committed sponsors** sign up to cover N students per cycle. A cycle is one UMass application season, September to January, named by the fall it leads to, so Sep 2026 to Jan 2027 is "Fall 2027". Pledges stay first come, first served. Committed sponsors get emails when the season opens and again in December, and also when a new applicant matches the research area or country they picked.
- **One application per person:** an email address or UMass applicant ID can only apply once.
- **Email verification:** after applying, the applicant gets a 6-digit code by email. Their card only appears on the board after they enter it, either straight away or later under Track my application.
- **Applicant confirms payment:** once sponsors have paid, the card shows "Paid · awaiting confirmation". The applicant confirms under Track my application when UMass shows the fee as paid. The card then leaves the board, is counted under "fees covered", and stays on the thank-you wall. If the applicant doesn't respond within 7 days, it's confirmed automatically.
- **Cover full or chip in:** "Cover the full $90" (or "Cover the last $X") offers **Pay now** or **Confirm pledge** (pay within 24 hours). "Chip in part" offers $10, $30, $50 or another amount below what's left.
- **Collector account (real version):** pool shares go into one organiser-run account, for example PayPal Business, and the organiser pays UMass with that account's card when a fee is complete. A full-fee sponsor can pay the same way.

## Files

- `index.html`: the whole website.
- `config.js`: where the Google Sheet link goes. If it's empty, the site runs in preview mode and each visitor's data stays in their own browser.
- `apps-script/Code.gs`: the Google Sheet backend.
- `data-umass-programs.json`: the UMass Amherst program list the Apply form uses.

## Connect the Google Sheet (about 5 minutes)

1. Create a new Google Sheet, for example "PhD Fee Bridge data".
2. Open **Extensions → Apps Script**. Delete what's in `Code.gs`, paste in all of `apps-script/Code.gs`, and save.
3. In the toolbar, pick the function **setup** and click **Run**, then approve the permissions. Google warns that the app is unverified because it's your own script, so choose Advanced → Go to project. This creates the tabs and a daily 9am job that releases expired pledges and sends reminders.
4. Optional: run **seedDemo** to add the fictional example data. Emails are never sent to `@example.com` addresses.
5. Click **Deploy → New deployment**, choose the type **Web app**, set *Execute as* to **Me** and *Who has access* to **Anyone**, then click **Deploy**.
6. Copy the Web app URL (it ends in `/exec`) and paste it into `config.js`:
   ```js
   window.FEE_BRIDGE_API = "https://script.google.com/macros/s/XXXX/exec";
   ```
7. Commit `config.js`. GitHub Pages updates within a minute or two.

If you change `Code.gs` later, paste the new version, run **setup** again (it adds any new columns and keeps your data), then go to **Deploy → Manage deployments → Edit → Version: New version** so the change goes live.

### Emails

The script sends emails from your Gmail: application numbers, pledge codes, "time to pay" notices, released pledges and season reminders. A free Gmail account can send about 100 a day. To turn emails off, set `SEND_EMAILS = false` at the top of `Code.gs`.

### What lands in the Sheet

- **Applications:** one row per applicant.
- **Pledges:** one row per pledge (amount, status `held` / `due` / `paid` / `released`, and the simulated payment).
- **Sponsors:** committed sponsors and which reminders they've had.
- **Log:** every event, with the time it happened.

Only you can open the Sheet. Emails, UMass applicant IDs, application numbers and codes never go to the public board. Anyone with the web app URL can use the same actions as the website: apply, pledge, pay and look up.

## Turn on GitHub Pages

In the repository, go to **Settings → Pages → Build and deployment**, set Source to "Deploy from a branch", choose Branch `main` and folder `/ (root)`, and save. The site appears at `https://kyatha.github.io/Fee_sponsorships/`.
