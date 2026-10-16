# SEBN TN2 — Reclamations system: complete go-live guide (≈ 25 min)

Folders in this package
* `backend/Code.gs` — ONE Google Apps Script that serves BOTH the form (verify + submit) and the HR app.
* `form/index.html`  — your reclamation form, identical to yours except ONE line: the backend address.
* `site/`            — the HR web app (index.html, config.js, assets).

You do NOT need your old form script any more: this new backend replaces it.

## STEP 1 — Google Sheet (the database)
1. Go to https://sheets.google.com → **Blank spreadsheet** → name it `SEBN TN2 – Réclamations`.
   (If you know the sheet your form was saving into and it holds real reclamations, use it instead.)
2. **File → Import → Upload** → choose `Reclamations_FAKE.xlsx` → "Insert new sheet(s)" → Import.
   Rename the new tab to exactly **Reclamations**.
   (Skip this if you want to start with an empty history: the backend creates the tab itself.)
3. **File → Import** again → `master_data_FAKE.xlsx` → "Insert new sheet(s)".
   Rename that tab to exactly **MasterData**. (It is the employee list used to check Matricule + CIN.)
4. Delete the empty default "Sheet1" tab. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/`**`THIS_PART`**`/edit`
   Keep this sheet PRIVATE (it contains CIN numbers). Do not share it with anyone but HR.

## STEP 2 — Backend (Google Apps Script)
1. Go to https://script.google.com → **New project** (a brand-new one — never reuse another project).
2. Delete the sample code, paste ALL of `backend/Code.gs`.
3. Replace `COLLER_ICI_L_ID_DU_GOOGLE_SHEET` (top of the file) with your Sheet ID. **Save** (Ctrl+S).
4. Choose the function `installer` in the toolbar → **Run**. Authorize: "Review permissions" → your account →
   "Advanced" → "Go to … (unsafe)" → Allow. (This warning is normal for your own script.)
   Check: the sheet now has tabs `RH_Dossiers` and `RH_Journal`; execution log says "Installation terminée".
5. **Deploy → New deployment** → ⚙ type **Web app** → Execute as: **Me** → Who has access: **Anyone** → Deploy.
   Copy the **Web app URL** (ends with `/exec`). Keep it: you need it twice.
6. Test it: paste the URL in a browser tab → you must see `{"ok":true,"service":"SEBN TN2 RH",...}`.

## STEP 3 — Reconnect the form
1. Open `form/index.html` in Notepad. Find `COLLER_ICI_L_URL_DU_NOUVEAU_BACKEND` (one place) and replace it with
   the Web app URL from step 2 (keep the quotes). Save.
2. Netlify → your form's site → **Deploys** tab → drag the `form` folder onto the "drag and drop" area.
   The form keeps the same address https://ornate-vacherin-e7ddbb.netlify.app/.
3. Test: open the form, enter a matricule + CIN from MasterData
   → you must see "Bonjour Ayari Habib" → fill the form → "Réclamation enregistrée avec succès".
   A new row appears in the `Reclamations` tab AND in `RH_Dossiers` with its ID.

## STEP 4 — Publish the HR app
1. Open `site/config.js`, paste the same Web app URL in `API_URL`. Save.
2. https://app.netlify.com/drop → drag the **`site` folder** (not `backend`, not `form`). Note the new link.
3. Open it, log in with 01756 / 04526 / 00600. Submit a test reclamation from the form: it appears on the
   dashboard within seconds. Take it, close it, check Historique and Journal d'activité.

## If something goes wrong
| Symptom | Cause / fix |
|---|---|
| Form says "Action inconnue" | The form still points to an old backend: redo step 3. |
| "Liste des employés introuvable" | The tab is not named exactly `MasterData`. |
| "Matricule ou CIN incorrect" for a valid employee | Check the employee exists in MasterData; 5 wrong tries block that matricule for 10 min. |
| "Connexion impossible" | Wrong/missing URL in `config.js` or form; deployment access not "Anyone". |
| You changed Code.gs but nothing changed | Deploy → Manage deployments → ✏ → Version: **New version** → Deploy. |
| New reclamations don't show | Run `installer` once; check Triggers (clock icon) has `surveiller` every minute. |
Never edit/delete rows in `RH_Dossiers` and `RH_Journal` by hand. To change agents: edit `AGENTS` in `Code.gs`, then a new version.
