# Security and Data Handling

## The script is read-only

`Get-PBIEnvironmentInventory.ps1` reads inventory metadata. It contains no code path that creates, modifies, or deletes anything in a Power BI or Fabric tenant.

Verify before you run it:

```powershell
.\Get-PBIEnvironmentInventory.ps1 -ListApiCalls
```

This prints every HTTP call the script can make, with the verb for each, and exits without authenticating or writing anything.

### Independent verification

Do not rely on the manifest alone — confirm it against the code.

```powershell
# 1. No write verbs. Returns nothing.
Select-String -Path .\Get-PBIEnvironmentInventory.ps1 -Pattern '-Method\s+(Put|Patch|Delete)'

# 2. Every HTTP call site, in context.
Select-String -Path .\Get-PBIEnvironmentInventory.ps1 -Pattern 'Invoke-PowerBIRestMethod|Invoke-RestMethod' -Context 0,1

# 3. Every command the script invokes, via the PowerShell parser.
$t=$null; $e=$null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    "$PWD\Get-PBIEnvironmentInventory.ps1", [ref]$t, [ref]$e)
$ast.FindAll({param($n) $n -is [System.Management.Automation.Language.CommandAst]}, $true) |
    ForEach-Object { $_.GetCommandName() } | Sort-Object -Unique
```

The only file-creating cmdlets are `New-Item` (creates the local `output` folder), `Out-File`, and `Export-Csv`. All three target the local `output` folder only.

### The single POST

The Scanner API `getInfo` endpoint is a POST because the list of workspace IDs to read is sent in the request body. It queues a **read-only metadata scan** and returns a scan ID. It modifies nothing. It runs only when `-DeepScan` is passed.

Reference: [Admin - WorkspaceInfo PostWorkspaceInfo](https://learn.microsoft.com/rest/api/power-bi/admin/workspace-info-post-workspace-info)

### Network-level proof

If code review isn't sufficient for your process, run the script through an HTTP proxy (Fiddler, mitmproxy) and inspect the verbs directly. You will see GETs, plus the one Scanner POST if `-DeepScan` is used.

---

## The output contains organizational data

The report is **not** anonymized. Treat it as internal.

It includes user principal names, display names, departments, email-style identifiers, workspace and report names, data source connection strings (server and database names), and gateway machine names.

Recommendations:

- Classify and label the output before distributing it.
- Do not commit `output/` to source control. It is gitignored by default.
- Connection strings can reveal internal infrastructure naming. Review the **Data Sources** and **Gateway Sources** tabs before sharing externally.
- The HTML report makes no network calls when opened. It is self-contained, with no CDN references, external scripts, or telemetry.

### Credentials

The script never stores, logs, or writes credentials. Authentication is delegated to `Connect-PowerBIServiceAccount` and the Azure CLI, which manage their own token caches. No token value is written into the JSON or HTML output.

---

## Reporting a problem

If you find a behavior that contradicts anything stated here, open an issue.
