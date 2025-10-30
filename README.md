# SharePointPS_Public

PowerShell utilities for SharePoint Online / OneDrive administration at scale.
Currently includes tooling to detect and remediate PUID mismatch issues that block access to OneDrive/SPO resources.

Repo: vipins5/SharePointPS_Public (branch: main). Files include PUIDMismatch.ps1. 

✨ What’s inside
PUIDMismatch.ps1 – Diagnose and (optionally) repair PUID mismatch for a specific user across SharePoint Online.
Typical symptoms: user can’t access their OneDrive or site after rename/restore/migration, odd “user not found” behavior, or permission ghosts.

File is present in the repo root. 

🧩 Why PUID mismatch?

Each M365 identity has a PUID baked into several places (AAD/Entra, SPO user info lists, OneDrive site collections, etc.).
When they drift (due to tenant moves, restores, profile corruption, or stale entries), SPO may resolve the wrong principal, causing access and permission anomalies. This script helps you:

-> Inspect relevant objects and IDs
-> Compare expected vs. actual PUIDs
-> Generate a detailed report (dry run by default)
-> Optionally execute safe remediation steps

✅ Features

Dry-run mode ($ReportMode = $true) to preview changes safely
App-only certificate auth (thumbprint or PFX)
Clear logging + CSV output for audit / RCA
Guardrails to avoid unintended changes (targeted to a single AffectedUser)

📦 Requirements

PowerShell 7+ (recommended)
Modules (latest stable):
PnP.PowerShell

Microsoft.Graph (or sub-modules where applicable)
Microsoft.Online.SharePoint.PowerShell (if your environment uses it for certain calls)
You’ll also need:
SPO Admin URL (e.g., https://contoso-admin.sharepoint.com)
Entra ID App Registration with certificate and the right permissions
Graph: User.Read.All, Sites.FullControl.All (as needed)
SharePoint: Sites.FullControl.All (app-only), or sites.selected + site grants for least privilege

🔐 Authentication options

The script supports both certificate approaches:
Thumbprint (certificate in CurrentUser\My)
PFX file path + password

**Tip:** Start in dry-run to validate findings. Flip $ReportMode = $false only after you review the report and understand the actions.
