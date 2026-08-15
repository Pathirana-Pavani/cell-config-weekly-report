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

1. Push this repo to GitHub (keep it **private** -- it's fine since it
   contains no real data, but no reason to make it public).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with
   GitHub.
3. Click "New app", pick this repo/branch, set the main file to `webapp.py`,
   click Deploy.
4. Share the resulting `https://<name>.streamlit.app` link with the team.

The app sleeps after inactivity and wakes up (a few seconds delay) on the
next visit -- normal for the free tier and fine for weekly use.

## What's NOT in this repo

`input_zip/`, `other_exports/`, and `output/` are gitignored -- they hold
real network configuration data and must never be committed. Only the
folder structure (`.gitkeep`) and the code are tracked.
