# Design Notes – Public Group Visibility Auto-Remediation

## Flow
1. Connect to Microsoft Graph (App Only)
2. List Unified groups → filter `Visibility = Public`
3. For each group:
   - Try resolve SiteUrl (PnP: `Get-PnPMicrosoft365Group -IncludeSiteUrl`)
   - Apply Sensitivity Label (site-level preferred; group-level fallback)
   - Log “ACTION: PUBLIC→PRIVATE …” and label id
4. Upload `logs_*.txt` and `report_*.txt` to SharePoint library
5. Power Automate watches for `log*.txt`, extracts 2 lines, posts to Teams

## Notes
- Sensitivity Label must enforce Private visibility policy
- Use app-only auth; no user context required
- Keep the Graph paging (`-All`) for large tenants
- Handle errors per group so one failure doesn’t stop the job
