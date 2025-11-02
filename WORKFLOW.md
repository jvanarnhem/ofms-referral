# OFMS Referral - GAS Sync Workflow

## Daily Workflow

### Making Changes in Google Apps Script
1. Edit code in Google Sheets (Extensions → Apps Script)
2. Test changes in the spreadsheet
3. Copy updated code from each file

### Updating GitHub
1. Navigate to repo: `cd /Users/jeff_v/ofms-referral`
2. Paste changes into corresponding files in `src/`
3. Check changes: `git status`
4. Add changes: `git add .`
5. Commit: `git commit -m "Update: [describe changes]"`
6. Push: `git push origin main`

## File Mapping
- Google Apps Script → Local Repository
  - Code.gs → src/Code.gs
  - Settings.gs → src/Settings.gs
  - admin.html → src/admin.html
  - forms2.html → src/forms2.html

## Last Sync
- Date: $(date +%Y-%m-%d)
- Version: 1.0.0
