# PhD Fee Bridge (demo)

Sponsors cover UMass Amherst graduate application fees ($90). Applicants post a short research statement; sponsors pledge, get a 24-hour hold and a pledge code, then pay. Payments in this demo are **simulated**. Nothing is charged and no card details are stored.

There are no logins. Applicants come back with their **application number + email**; sponsors with their **pledge code + email**.

Not affiliated with or endorsed by UMass Amherst.

## Files

- `index.html`: the whole website.
- `config.js`: where the Google Sheet link goes. Empty means preview mode (each visitor's data stays in their own browser).
- `apps-script/Code.gs`: the Google Sheet backend.

## Connect the Google Sheet (about 5 minutes)

1. Create a new Google Sheet (for example "PhD Fee Bridge data").
2. Open **Extensions → Apps Script**. Delete what's in `Code.gs`, paste in all of `apps-script/Code.gs`, and save.
3. In the toolbar, pick the function **setup** and click **Run**. Approve the permissions (Google warns that the app is unverified, because it's your own script. Choose Advanced → Go to project).
4. Optional: run **seedDemo** to add the five fictional example applicants.
5. Click **Deploy → New deployment**, choose the type **Web app**, set *Execute as* to **Me** and *Who has access* to **Anyone**, then click **Deploy**.
6. Copy the Web app URL (it ends in `/exec`) and paste it into `config.js`:
   ```js
   window.FEE_BRIDGE_API = "https://script.google.com/macros/s/XXXX/exec";
   ```
7. Commit `config.js`. GitHub Pages updates within a minute or two.

If you change `Code.gs` later, go to **Deploy → Manage deployments → Edit → Version: New version** so the change goes live.

### What lands in the Sheet

- **Applications** tab: one row per applicant, with their status, pledge and simulated payment. Email, UMass applicant ID, application number and pledge code stay in the Sheet. The website only ever receives the public fields.
- **Log** tab: every application, pledge, payment and release, with the time it happened.

Only you can open the Sheet. Anyone with the web app URL can use the site's actions (apply, pledge, pay, look up), the same as visiting the site.

## Turn on GitHub Pages

Repository **Settings → Pages → Build and deployment**: Source "Deploy from a branch", Branch `main`, folder `/ (root)`. The site appears at `https://<username>.github.io/<repo>/`.
