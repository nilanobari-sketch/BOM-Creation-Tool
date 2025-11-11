# BOM Creation Tool — One-File EXE via GitHub Actions

**How to build**
1. Make sure your Python script is at the repo root: `BOM Creation Tool_V8 (1).py`
2. The workflow file should exist at: `.github/workflows/build.yml`
3. Go to the **Actions** tab and run the **Build BOM Creation Tool EXE** workflow.

**How to download the EXE**
- When the workflow finishes, open the run page.
- In the **Artifacts** section, download `BOM-Creation-Tool-EXE.zip`.
- Inside, you’ll find: `dist\BOM Creation Tool.exe`

**Optional: Code signing**
- Add GitHub secrets:
  - `CODE_SIGN_PFX` = Base64 of your `.pfx`
  - `CODE_SIGN_PFX_PASSWORD` = certificate password
- Re-run the workflow to get a signed EXE.
