# Weekly Cell Config Report

Fills the weekly `EUtranCellFDD` report template from the network export zip,
plus two optional side-lookups (Cell DB traffic, power license). Two ways to
run it:

## Command line (original workflow)

```
input_zip/        <- drop the weekly export zip here
other_exports/    <- optionally drop "Cell DB Export*.xlsx" and/or "powerlicense*.xlsx" here
template/         <- basic_configurations.xlsx lives here (already committed)
output/           <- filled result appears here
```

```
pip install -r requirements.txt
python excel_automation.py
```

## Web interface (for the team)

A Streamlit app (`webapp.py`) wraps the same logic with file uploads instead
of folders, so anyone can generate the report from a link.

Run locally:

```
pip install -r requirements.txt
streamlit run webapp.py
```

Deploy for free (one-time setup):

1. Push this repo to GitHub and keep it **private** -- required, since
   (per below) generated reports containing real network data get
   committed into it.
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub.
3. Click "New app", pick this repo/branch, set the main file to `webapp.py`,
   click Deploy.
4. Share the resulting `https://<name>.streamlit.app` link with the team.

The app sleeps after inactivity and wakes up (a few seconds delay) on the
next visit -- normal for the free tier and fine for weekly use.

### Persistent report history (optional but recommended)

The "Previous Reports" tab commits each day's report straight into this
repo (`reports/<date>.xlsx`, overwritten if regenerated the same day), or
just shows a setup message if unconfigured (report generation/download
still work either way). This was chosen after Cloudflare R2, Supabase, and
Firebase Storage all turned out to require a billing card just to enable
their storage product -- committing to this repo needs no new account and
no card, at the cost of real network data living in git history (hence
the repo must stay private).

One-time setup:

1. Create a GitHub Personal Access Token (fine-grained, scoped to just
   this repo, "Contents: Read and write" permission).
2. In the Streamlit Cloud app's Settings -> Secrets, add:

   ```toml
   [github]
   token = "github_pat_..."
   ```

## What's NOT in this repo

`input_zip/` and `other_exports/` are gitignored -- they hold the *input*
network configuration data and must never be committed. `output/` (the
CLI's local output folder) is also gitignored. The `reports/` folder,
however, IS committed -- see "Persistent report history" above.
