# Power BI Environment Inventory

A single PowerShell script that produces a complete, factual inventory of a Power BI / Fabric tenant.

It answers one question: **what is actually in this environment?**

It does not score your tenant, rank your workspaces, grade your governance, or tell you what to buy. Every number in the output comes back from a Microsoft API or is a direct count of API values, so anything in the report can be traced to a service response and re-verified.

---

## Read-only. Verifiably.

This script **reads inventory data and nothing else.** It does not create, modify, or delete workspaces, reports, datasets, permissions, capacities, gateways, users, or tenant settings. There is no code path that can.

You do not have to take that on faith. Print every HTTP call the script is capable of making, without signing in to anything:

```powershell
.\Get-PBIEnvironmentInventory.ps1 -ListApiCalls
```

That prints the manifest below and exits immediately — no authentication, no collection, no output files.

| Verb | API | Endpoint | When |
| --- | --- | --- | --- |
| GET | Power BI Admin | `/admin/capacities` | Always |
| GET | Power BI Admin | `/admin/capacities/refreshables` | Always |
| GET | Power BI Admin | `/admin/groups` | Always |
| GET | Power BI Admin | `/admin/datasets/{id}/datasources` | Always |
| GET | Power BI Admin | `/admin/datasets/{id}/refreshSchedule` | Always |
| GET | Power BI Admin | `/admin/gateways` | Always |
| GET | Power BI Admin | `/admin/gateways/{id}/datasources` | Always |
| GET | Power BI Admin | `/admin/pipelines` | Always |
| GET | Power BI Admin | `/admin/apps` | Always |
| GET | Fabric Admin | `/v1/admin/tenantsettings` | Always (needs Fabric Admin role) |
| GET | Microsoft Graph | `/v1.0/users` | Unless `-SkipGraph` |
| POST | Power BI Admin | `/admin/workspaces/getInfo` | Only with `-DeepScan` |
| GET | Power BI Admin | `/admin/workspaces/scanStatus/{id}` | Only with `-DeepScan` |
| GET | Power BI Admin | `/admin/workspaces/scanResult/{id}` | Only with `-DeepScan` |

**About the one POST.** Thirteen of the fourteen calls are GETs. The exception is the Power BI Scanner API's `getInfo` endpoint, which is a POST *because of how the API is designed*, not because anything is being written: the list of workspace IDs to read is too long for a query string, so it travels in the request body. The call queues a read-only metadata scan and returns a scan ID you then poll. It writes nothing to the tenant, and it only runs if you explicitly pass `-DeepScan`. See the [Microsoft documentation for Admin - WorkspaceInfo PostWorkspaceInfo](https://learn.microsoft.com/rest/api/power-bi/admin/workspace-info-post-workspace-info).

**There is no PUT, PATCH, or DELETE call anywhere in the script.** You can confirm that yourself:

```powershell
# Should return nothing.
Select-String -Path .\Get-PBIEnvironmentInventory.ps1 -Pattern '-Method\s+(Put|Patch|Delete)'

# Lists every HTTP call site so you can read them in context.
Select-String -Path .\Get-PBIEnvironmentInventory.ps1 -Pattern 'Invoke-PowerBIRestMethod|Invoke-RestMethod'
```

Or audit every command the script invokes, using PowerShell's own parser rather than a text search:

```powershell
$t=$null; $e=$null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    "$PWD\Get-PBIEnvironmentInventory.ps1", [ref]$t, [ref]$e)
$ast.FindAll({param($n) $n -is [System.Management.Automation.Language.CommandAst]}, $true) |
    ForEach-Object { $_.GetCommandName() } | Sort-Object -Unique
```

You will not find a single tenant-modifying cmdlet in the result. The only `New-Item` call creates the local `output` folder.

### What it *does* write

| Location | What | Why |
| --- | --- | --- |
| `.\output\*.json` / `*.html` | The report | The deliverable |
| `.\output\csv\*.csv` | Flat exports | Only with `-ExportCsv` |
| Local token cache | Sign-in state | Standard `Connect-PowerBIServiceAccount` and `az login` behavior |

Nothing else on your machine is touched, and nothing at all in the tenant.

### Permissions it needs

Read-only administrative visibility: **Fabric Administrator** or **Power BI Administrator**. These roles can also change things — the script simply never does. If you want to prove that at the network layer rather than the code layer, run it behind Fiddler or the browser dev tools and watch the verbs.

---

## What it collects

| Area | Detail |
| --- | --- |
| **Capacities** | SKU, region, state, admins |
| **Workspaces** | Capacity assignment, license mode, description, admin list, item counts, workspaces with no content, workspaces with no admin |
| **Access** | Every access entry with role and principal type — users, groups, and service principals |
| **Users** | Per-user role rollup across the whole tenant, guest flag, PPU workspace count |
| **M365 licenses** | Who holds Pro / PPU / Fabric Free, whether they appear in any workspace ACL |
| **Semantic models** | Storage mode, refreshability, RLS requirement, gateway requirement, owner, created date |
| **Data sources** | Every dataset-to-source binding, with connection detail and gateway ID |
| **Refresh** | Configured schedules, plus actual refresh counts, failure counts, durations, and last outcome |
| **Content** | Reports, paginated reports, dashboards, dataflows |
| **Gateways** | On-premises, VNet, and personal gateways with their data sources |
| **Pipelines** | Deployment pipelines and their workspace stages |
| **Apps** | Published apps |
| **Tenant settings** | Setting switches and the security groups scoped to them |
| **Deep scan** *(optional)* | Endorsement (certified / promoted) and sensitivity labels |

---

## Requirements

- **PowerShell 7+** — [aka.ms/powershell](https://aka.ms/powershell). Do **not** use PowerShell ISE; it is frozen at 5.1 and will misbehave.
- **MicrosoftPowerBIMgmt** module:
  ```powershell
  Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser
  ```
- **Fabric Administrator** or **Power BI Administrator** role.
- *Optional:* **Azure CLI** for the Microsoft 365 license section — [aka.ms/installazurecliwindows](https://aka.ms/installazurecliwindows), then `az login --tenant <tenant-id>`. Skip it with `-SkipGraph` and the rest still runs.

---

## Usage

```powershell
# Unblock if downloaded
Unblock-File .\Get-PBIEnvironmentInventory.ps1

# Review what it will call, before running it for real
.\Get-PBIEnvironmentInventory.ps1 -ListApiCalls

# Interactive - prompts for tenant and options
.\Get-PBIEnvironmentInventory.ps1

# Unattended
.\Get-PBIEnvironmentInventory.ps1 -TenantId <guid> -ExportCsv -SkipGraph

# Everything, including endorsement and sensitivity labels
.\Get-PBIEnvironmentInventory.ps1 -TenantId <guid> -DeepScan -ExportCsv
```

### Parameters

| Parameter | Effect |
| --- | --- |
| `-TenantId` | Directory (tenant) ID. Prompted for if omitted. |
| `-SkipGraph` | Skip the Microsoft 365 license collection. |
| `-DeepScan` | Add endorsement and sensitivity labels via the Scanner API. Slower on large tenants. |
| `-ExportCsv` | Also write one CSV per collection. |
| `-ListApiCalls` | Print the API manifest and exit. Connects to nothing. |

Your tenant ID is in **Azure Portal → Microsoft Entra ID → Overview**.

---

## Output

Written to `.\output\`:

| File | Contents |
| --- | --- |
| `PBI_Environment_Inventory.html` | Browsable report — summary counts, breakdown tables, and 18 searchable/sortable tabs with per-tab CSV download |
| `PBI_Environment_Inventory.json` | The same data, structured, for pipelines or further analysis |
| `csv\*.csv` | One flat file per collection (with `-ExportCsv`) |

The HTML file is fully self-contained — no CDN, no external scripts, no telemetry, no network calls when opened. Mail it or drop it in SharePoint and it works.

---

## Notes on accuracy

- **Sections that fail are reported, not hidden.** If your account can't read tenant settings, the report says so in a banner at the top instead of quietly showing zero.
- **Throttling is handled.** HTTP 429 responses back off exponentially and retry.
- **Counts are traceable.** Summary tiles are counts of the rows in the tabs below them. If a tile says 412 semantic models, the Semantic Models tab has 412 rows.
- **`My Workspace` is included but separated.** Personal workspaces are counted apart from shared workspaces, because mixing them inflates workspace counts.
- **No interpretation.** Columns like `HasNoContent` and `HasWorkspaceAdmin` are facts (`TotalItemCount = 0`, `AdminCount = 0`), not judgments. What they mean for your environment is your call.

---

## Credit

Structure and API approach adapted from the [Power BI Premium to Fabric assessment script](https://github.com/Diaz506/powerbi-premium-to-fabric-assessment) by Steven Uba. This version drops the migration and cost-modeling layer in favor of a neutral, customer-facing environment census.

## License

MIT — see [LICENSE](LICENSE).
