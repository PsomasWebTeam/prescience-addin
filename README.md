# Prescience by Psomas — the "Send to Prescience" button in Outlook

The static web page behind the **Send to Prescience** button in Outlook (classic and new) for
the Psomas Legal Department, served by GitHub Pages at
`https://psomaswebteam.github.io/prescience-addin/`.

- `manifest.xml` — what each person adds to their Outlook once (Get Add-ins → My add-ins →
  Add a custom add-in → Add from file), or what IT uploads in the Microsoft 365 admin center
  for everyone.
- `taskpane.html` / `taskpane.js` / `taskpane.css` — the side panel: shows the selected email,
  **File this email** or **Compare an attachment** (with a pick list when the email carries
  several documents), sends the file to the "Prescience Mail" Power Automate flow, which puts
  it in the shared inbox folder that Prescience on the person's PC watches.
- `assets/` — the button icons.

This site stores nothing. The flow's web address is typed in once per user and kept in
Outlook's roaming settings for that user; no secret lives in these files.

Published from the private repository PsomasWebTeam/Prescience (`addin/` folder) by
`deploy_addin.py`. It is a SEPARATE add-in from the Contract Queue button (its own manifest,
its own site, its own flow) — Matt's decision of 2026-10-10.
