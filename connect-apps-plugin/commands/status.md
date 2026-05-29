---
description: Show the connection status across all supported apps (Azure, Microsoft 365, GitHub)
allowed-tools: Bash, Read
---

Report a concise connection-status summary for every app this plugin supports.
Run each check defensively — a missing tool or absent session is a normal
"not connected" result, not an error to halt on.

Check each, then print a single table (App | Tool | Connected as | Notes):

1. **Azure**
   - CLI: `az account show -o json` (parse `user.name` / `name`)
   - PowerShell: `Get-AzContext`
2. **Microsoft 365 / Graph**
   - PowerShell: `Get-MgContext` (Graph), `Get-ConnectionInformation`
     (Exchange Online) if those modules are loaded
3. **GitHub**
   - `gh auth status` and `gh api user --jq .login`, else `git remote -v`

For each app show whether it is connected and which identity/tenant is active.
End with a short note on which apps still need `/connect-apps-plugin:<app>`
to connect.
