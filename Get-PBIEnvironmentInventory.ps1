#Requires -Version 5.1
<#
.SYNOPSIS
    Power BI Environment Inventory — full tenant census. READ-ONLY.

.DESCRIPTION
    Connects to the Power BI Admin REST API (and optionally Microsoft Graph and the
    Fabric Admin API) and exports a complete factual inventory of a Power BI tenant.

    READ-ONLY BY DESIGN
    This script never creates, modifies or deletes tenant content, permissions,
    capacities or settings. It calls inventory endpoints only. Run it with
    -ListApiCalls to print every HTTP call it is capable of making, with the verb
    for each, and exit without connecting. The only non-GET call is the Scanner API
    'getInfo' POST, which is a POST purely because the workspace id list travels in
    the request body; it queues a read-only metadata scan and is used only with
    -DeepScan. The only writes are local output files.

    This script reports WHAT EXISTS. It does not score, rank, grade, or recommend.
    Every column is either returned directly by the API or is a direct count/derivation
    of API values, so any number in the report can be traced back to a service response.

    COLLECTED:
      - Capacities (SKU, region, state, admins)
      - Workspaces (capacity, license mode, content counts, description)
      - Workspace access entries (users, groups, service principals, with role)
      - Per-user role rollup across the tenant
      - Microsoft 365 users holding a Power BI/Fabric license (optional, needs Graph)
      - Semantic models / datasets (storage mode, RLS, owner, created date)
      - Dataset data sources and gateway bindings
      - Scheduled refresh configuration
      - Refreshable items with recent refresh counts and last refresh outcome
      - Reports (Power BI and paginated), dashboards, dataflows
      - Gateways (on-premises, VNet, personal) and their data sources
      - Deployment pipelines and stages
      - Published apps
      - Tenant settings (optional, needs Fabric Admin API access)
      - Endorsement and sensitivity labels (optional deep scan via Scanner API)

    OUTPUT (./output):
      PBI_Environment_Inventory.json   Structured data, one node per collection
      PBI_Environment_Inventory.html   Tabbed browsable report with search + CSV export
      csv\*.csv                        One flat CSV per collection (optional)

.REQUIREMENTS
    - MicrosoftPowerBIMgmt module   Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser
    - Fabric Administrator or Power BI Administrator role
    - PowerShell 7 recommended. Do not use PowerShell ISE.
    - Optional: Azure CLI (az login) for the Microsoft 365 license section

.PARAMETER TenantId
    Directory (tenant) ID. Prompted for if omitted.

.PARAMETER SkipGraph
    Skip the Microsoft 365 licensed-user collection.

.PARAMETER DeepScan
    Run the Scanner API pass to add endorsement, sensitivity labels and detailed
    lineage. Slower on large tenants.

.PARAMETER ExportCsv
    Also write one CSV per collection into output\csv.

.PARAMETER ListApiCalls
    Print the full read-only API manifest (every endpoint and HTTP verb the script
    can call) and exit. Connects to nothing. Use this for security review.

.EXAMPLE
    .\Get-PBIEnvironmentInventory.ps1 -ListApiCalls
    Shows every call the script can make, without signing in or collecting anything.

.EXAMPLE
    .\Get-PBIEnvironmentInventory.ps1 -TenantId <guid> -ExportCsv
    Standard inventory run with CSV export.

.NOTES
    READ-ONLY. The script issues GET requests, plus the single Scanner API POST that
    queues a read-only metadata scan when -DeepScan is used. It contains no PUT, PATCH
    or DELETE call and no tenant-modifying cmdlet. It never changes tenant content or
    settings. Local writes are limited to the .\output folder.
    API reference: https://learn.microsoft.com/rest/api/power-bi/admin
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [switch]$SkipGraph,
    [switch]$DeepScan,
    [switch]$ExportCsv,
    [switch]$ListApiCalls
)

# ============================================================
# READ-ONLY GUARANTEE
# ------------------------------------------------------------
# This script does not create, modify, or delete anything in the
# Power BI / Fabric tenant. It calls only inventory endpoints.
#
# Every HTTP call the script can make is declared in $ApiManifest
# below. Run with -ListApiCalls to print that manifest and exit
# without connecting or collecting anything.
#
# Of those calls, exactly one is not an HTTP GET: the Scanner API
# 'getInfo' endpoint is a POST *by API design* because the request
# body carries the list of workspace IDs to read. It queues a
# read-only metadata scan and returns a scan id. It writes nothing
# to the tenant. It is only called with -DeepScan.
#
# The only things the script writes are local files in .\output,
# plus token caches created by the standard sign-in flows.
# ============================================================
$ApiManifest = @(
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/capacities";                        Purpose="Capacity list, SKU, region, admins";      Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/capacities/refreshables";           Purpose="Refresh counts, failures, last outcome";  Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/groups";                            Purpose="Workspaces + users/reports/datasets";     Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/datasets/{id}/datasources";         Purpose="Data source bindings per model";          Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/datasets/{id}/refreshSchedule";     Purpose="Configured refresh windows";              Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/gateways";                          Purpose="Gateway inventory";                       Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/gateways/{id}/datasources";         Purpose="Data sources behind each gateway";        Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/pipelines";                         Purpose="Deployment pipelines and stages";         Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/apps";                              Purpose="Published apps";                          Scope="Always" }
    [PSCustomObject]@{ Verb="GET";  Api="Fabric Admin";    Endpoint="/v1/admin/tenantsettings";                 Purpose="Tenant setting switches";                 Scope="Always (optional perms)" }
    [PSCustomObject]@{ Verb="GET";  Api="Microsoft Graph"; Endpoint="/v1.0/users";                              Purpose="Power BI / Fabric license assignment";    Scope="Unless -SkipGraph" }
    [PSCustomObject]@{ Verb="POST"; Api="Power BI Admin";  Endpoint="/admin/workspaces/getInfo";                Purpose="Queue read-only metadata scan (body = workspace ids)"; Scope="Only with -DeepScan" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/workspaces/scanStatus/{id}";        Purpose="Poll scan job state";                     Scope="Only with -DeepScan" }
    [PSCustomObject]@{ Verb="GET";  Api="Power BI Admin";  Endpoint="/admin/workspaces/scanResult/{id}";        Purpose="Endorsement + sensitivity labels";        Scope="Only with -DeepScan" }
)

function Show-ApiManifest {
    Write-Host ""
    Write-Host "=========================================================" -ForegroundColor Cyan
    Write-Host "  READ-ONLY API MANIFEST" -ForegroundColor Cyan
    Write-Host "  Every call this script is capable of making" -ForegroundColor Cyan
    Write-Host "=========================================================" -ForegroundColor Cyan
    $ApiManifest | Format-Table Verb, Api, Endpoint, Scope, Purpose -AutoSize -Wrap
    Write-Host "  GET calls  : $(@($ApiManifest | Where-Object { $_.Verb -eq 'GET' }).Count)" -ForegroundColor Green
    Write-Host "  POST calls : $(@($ApiManifest | Where-Object { $_.Verb -eq 'POST' }).Count)  (Scanner getInfo - queues a read-only scan, writes nothing)" -ForegroundColor Yellow
    Write-Host "  No PUT, PATCH or DELETE call exists anywhere in this script." -ForegroundColor Green
    Write-Host ""
    Write-Host "  Local writes: .\output\*.json, *.html, optional csv\*.csv" -ForegroundColor DarkGray
    Write-Host "  Tenant writes: none." -ForegroundColor DarkGray
    Write-Host ""
}

if ($ListApiCalls) { Show-ApiManifest; return }

# ============================================================
# SCRIPT SETTINGS
# ============================================================
$ScriptRoot   = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$OutputFolder = Join-Path $ScriptRoot "output"
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

$BatchSize   = 100      # page size for admin list endpoints
$DelayMs     = 200      # base pause between calls, keeps us under the throttle
$MaxRetry    = 5
$Script:Warnings = @()

function Write-Step { param([string]$Msg) Write-Host "$Msg" -ForegroundColor Yellow }
function Write-Ok   { param([string]$Msg) Write-Host "      $Msg" -ForegroundColor Green }
function Write-Dim  { param([string]$Msg) Write-Host "      $Msg" -ForegroundColor DarkGray }
function Add-Warn {
    param([string]$Area, [string]$Message)
    $Script:Warnings += [PSCustomObject]@{ Area = $Area; Message = $Message }
    Write-Host "      [SKIPPED] $Area : $Message" -ForegroundColor DarkYellow
}

# ============================================================
# SETUP
# ============================================================
function Read-Config {
    Write-Host ""
    Write-Host "=========================================================" -ForegroundColor Cyan
    Write-Host "  POWER BI ENVIRONMENT INVENTORY" -ForegroundColor Cyan
    Write-Host "  Full tenant census - no scoring, no recommendations" -ForegroundColor Cyan
    Write-Host "=========================================================" -ForegroundColor Cyan
    Write-Host "  READ-ONLY: this script only reads inventory data." -ForegroundColor Green
    Write-Host "  Nothing in your tenant is created, changed or deleted." -ForegroundColor Green
    Write-Host "  Run with -ListApiCalls to see every call it can make." -ForegroundColor DarkGray
    Write-Host ""

    if (-not $TenantId) {
        do {
            $TenantId = (Read-Host "  Tenant ID (Azure Portal > Microsoft Entra ID > Overview)").Trim()
            if ($TenantId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') {
                Write-Host "  Not a valid GUID. Try again." -ForegroundColor Red
                $TenantId = $null
            }
        } while (-not $TenantId)
    }

    if (-not $PSBoundParameters.ContainsKey('DeepScan')) {
        $Ans = (Read-Host "  Run deep scan for endorsement + sensitivity labels? (Y/N)").Trim()
        $Script:DoDeepScan = ($Ans -match '^[Yy]')
    } else { $Script:DoDeepScan = [bool]$DeepScan }

    if (-not $PSBoundParameters.ContainsKey('SkipGraph')) {
        $Ans = (Read-Host "  Include Microsoft 365 license data via Graph? (Y/N)").Trim()
        $Script:DoGraph = ($Ans -match '^[Yy]')
    } else { $Script:DoGraph = -not [bool]$SkipGraph }

    if (-not $PSBoundParameters.ContainsKey('ExportCsv')) {
        $Ans = (Read-Host "  Also export CSV files? (Y/N)").Trim()
        $Script:DoCsv = ($Ans -match '^[Yy]')
    } else { $Script:DoCsv = [bool]$ExportCsv }

    $Script:TenantId = $TenantId
    Write-Host ""
}

# ============================================================
# AUTH + API PLUMBING
# ============================================================
function Connect-PBIService {
    param([string]$TenantId)
    Write-Host "[AUTH] Connecting to Power BI Service..." -ForegroundColor Cyan
    if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt.Profile)) {
        throw "MicrosoftPowerBIMgmt is not installed. Run: Install-Module MicrosoftPowerBIMgmt -Scope CurrentUser"
    }
    $WarningPreference = "SilentlyContinue"
    $P = @{}
    if ($TenantId) { $P["Tenant"] = $TenantId }
    Connect-PowerBIServiceAccount @P | Out-Null
    $WarningPreference = "Continue"
    Write-Host "[AUTH] Connected.`n" -ForegroundColor Green
}

# Single GET with 429-aware retry.
function Invoke-PBIGet {
    param([string]$Url, [switch]$Quiet)
    $Attempt = 0
    while ($Attempt -lt $MaxRetry) {
        try {
            return (Invoke-PowerBIRestMethod -Url $Url -Method Get | ConvertFrom-Json)
        } catch {
            $Err = $_.ToString()
            if ($Err -like "*429*" -or $Err -like "*TooManyRequests*") {
                $Attempt++
                $Wait = [math]::Pow(2, $Attempt) * 1000
                Write-Host "      [429] backing off $([int]($Wait/1000))s (retry $Attempt/$MaxRetry)" -ForegroundColor Yellow
                Start-Sleep -Milliseconds $Wait
            } else {
                if (-not $Quiet) { throw }
                return $null
            }
        }
    }
    if (-not $Quiet) { throw "Throttled after $MaxRetry retries: $Url" }
    return $null
}

function Get-ValueArray {
    param($Response)
    if ($null -eq $Response) { return @() }
    if ($Response.PSObject.Properties['value']) { return @($Response.value) }
    return @($Response)
}

# Paged GET against an admin list endpoint using $top/$skip.
function Invoke-PBIAdminList {
    param([string]$Endpoint, [hashtable]$QueryParams = @{})
    $BaseUrl = "https://api.powerbi.com/v1.0/myorg/admin/$Endpoint"
    $Results = @()
    $Skip    = 0
    do {
        $QueryParams['$top']  = $BatchSize
        $QueryParams['$skip'] = $Skip
        $QS  = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"
        $Url = "$BaseUrl`?$QS"

        $Response = Invoke-PBIGet -Url $Url -Quiet
        $Items    = Get-ValueArray $Response
        if (-not $Items -or $Items.Count -eq 0) { break }

        $Results += $Items
        $Skip    += $BatchSize
        Start-Sleep -Milliseconds $DelayMs
    } while ($Items.Count -eq $BatchSize)
    return $Results
}

function Get-PrincipalName {
    param($Entry)
    if ($Entry.emailAddress)  { return $Entry.emailAddress }
    if ($Entry.identifier)    { return $Entry.identifier }
    if ($Entry.displayName)   { return $Entry.displayName }
    if ($Entry.graphId)       { return $Entry.graphId }
    return ""
}

function ConvertTo-IsoDate {
    param($Value)
    if (-not $Value) { return "" }
    try { return ([datetime]$Value).ToString("yyyy-MM-dd HH:mm") } catch { return [string]$Value }
}

# ============================================================
# COLLECTORS
# ============================================================

function Get-Capacities {
    Write-Step "[1/12] Capacities"
    $Rows = @()
    try {
        $Raw = Get-ValueArray (Invoke-PBIGet -Url "https://api.powerbi.com/v1.0/myorg/admin/capacities")
        foreach ($C in $Raw) {
            $Rows += [PSCustomObject]@{
                CapacityId   = $C.id
                CapacityName = $C.displayName
                Sku          = $C.sku
                Region       = $C.region
                State        = $C.state
                AdminCount   = if ($C.admins) { @($C.admins).Count } else { 0 }
                Admins       = if ($C.admins) { ($C.admins -join "; ") } else { "" }
                AccessRight  = $C.capacityUserAccessRight
            }
        }
        Write-Ok "$($Rows.Count) capacities."
    } catch { Add-Warn "Capacities" $_.Exception.Message }
    return $Rows
}

function Get-AllWorkspacesRaw {
    Write-Step "[2/12] Workspaces (expanded)"
    $All = Invoke-PBIAdminList -Endpoint "groups" -QueryParams @{
        '$expand' = 'users,reports,datasets,dashboards,dataflows'
    }
    Write-Dim "Returned by API      : $($All.Count)"
    $Active = @($All | Where-Object { $_.state -eq "Active" })
    Write-Ok "Active workspaces    : $($Active.Count)"
    Write-Dim "Personal (My Workspace): $(@($Active | Where-Object { $_.type -eq 'PersonalGroup' }).Count)"
    return $All
}

function Build-WorkspaceInventory {
    param($Workspaces, $CapacityRows)
    Write-Step "[3/12] Workspace + access inventory"

    $CapMap = @{}
    foreach ($C in $CapacityRows) { $CapMap[$C.CapacityId] = $C }

    $WsRows     = @()
    $AccessRows = @()
    $UserAgg    = @{}
    $SeenUpns   = @{}

    $i = 0
    foreach ($Ws in $Workspaces) {
        $i++
        if ($i % 500 -eq 0) { Write-Dim "workspace $i / $($Workspaces.Count)" }

        $Cap = if ($Ws.capacityId) { $CapMap[$Ws.capacityId] } else { $null }
        $CapName = if ($Cap) { $Cap.CapacityName }
                   elseif ($Ws.capacityId) { "Unresolved ($($Ws.capacityId))" }
                   else { "" }

        $Reports    = @($Ws.reports)
        $PbiReports = @($Reports | Where-Object { $_.reportType -ne "PaginatedReport" })
        $PagReports = @($Reports | Where-Object { $_.reportType -eq "PaginatedReport" })
        $Datasets   = @($Ws.datasets)
        $Dashboards = @($Ws.dashboards)
        $Dataflows  = @($Ws.dataflows)
        $Users      = @($Ws.users)

        $Admins = @($Users | Where-Object { $_.groupUserAccessRight -eq "Admin" })
        $ItemTotal = $PbiReports.Count + $PagReports.Count + $Datasets.Count + $Dashboards.Count + $Dataflows.Count

        $WsRows += [PSCustomObject]@{
            WorkspaceId       = $Ws.id
            WorkspaceName     = $Ws.name
            Type              = $Ws.type
            State             = $Ws.state
            LicenseMode       = if ($Ws.licenseType) { $Ws.licenseType } else { "" }
            OnDedicatedCapacity = [bool]$Ws.isOnDedicatedCapacity
            CapacityId        = $Ws.capacityId
            CapacityName      = $CapName
            CapacitySku       = if ($Cap) { $Cap.Sku } else { "" }
            CapacityRegion    = if ($Cap) { $Cap.Region } else { "" }
            Description       = $Ws.description
            HasWorkspaceAdmin = ($Admins.Count -gt 0)
            AdminCount        = $Admins.Count
            Admins            = (($Admins | ForEach-Object { Get-PrincipalName $_ }) -join "; ")
            AccessEntryCount  = $Users.Count
            ReportCount       = $PbiReports.Count
            PaginatedReportCount = $PagReports.Count
            DatasetCount      = $Datasets.Count
            DashboardCount    = $Dashboards.Count
            DataflowCount     = $Dataflows.Count
            TotalItemCount    = $ItemTotal
            HasNoContent      = ($ItemTotal -eq 0)
            PipelineId        = $Ws.pipelineId
        }

        foreach ($U in $Users) {
            $Name = Get-PrincipalName $U
            if (-not $Name) { continue }
            $PType = if ($U.principalType) { $U.principalType } else { "User" }

            $AccessRows += [PSCustomObject]@{
                WorkspaceId   = $Ws.id
                WorkspaceName = $Ws.name
                CapacityName  = $CapName
                PrincipalName = $Name
                DisplayName   = $U.displayName
                PrincipalType = $PType
                Role          = $U.groupUserAccessRight
                UserType      = $U.userType
                IsGuest       = ($Name -like "*#EXT#*")
            }

            if ($PType -ne "User") { continue }
            if ($Name -notlike "*@*") { continue }

            $Key = $Name.ToLower()
            $SeenUpns[$Key] = $true
            if (-not $UserAgg.ContainsKey($Key)) {
                $UserAgg[$Key] = [PSCustomObject]@{
                    UserPrincipalName = $Name
                    DisplayName       = if ($U.displayName) { $U.displayName } else { $Name }
                    UserType          = $U.userType
                    IsGuest           = ($Name -like "*#EXT#*")
                    WorkspaceCount    = 0
                    AdminRoles        = 0
                    MemberRoles       = 0
                    ContributorRoles  = 0
                    ViewerRoles       = 0
                    PPUWorkspaceCount = 0
                    Roles             = ""
                }
            }
            $A = $UserAgg[$Key]
            $A.WorkspaceCount++
            if ($Ws.licenseType -eq "PremiumPerUser") { $A.PPUWorkspaceCount++ }
            switch ($U.groupUserAccessRight) {
                "Admin"       { $A.AdminRoles++ }
                "Member"      { $A.MemberRoles++ }
                "Contributor" { $A.ContributorRoles++ }
                "Viewer"      { $A.ViewerRoles++ }
            }
        }
    }

    foreach ($K in $UserAgg.Keys) {
        $A = $UserAgg[$K]
        $Parts = @()
        if ($A.AdminRoles)       { $Parts += "Admin x$($A.AdminRoles)" }
        if ($A.MemberRoles)      { $Parts += "Member x$($A.MemberRoles)" }
        if ($A.ContributorRoles) { $Parts += "Contributor x$($A.ContributorRoles)" }
        if ($A.ViewerRoles)      { $Parts += "Viewer x$($A.ViewerRoles)" }
        $A.Roles = $Parts -join ", "
    }

    Write-Ok "Workspaces: $($WsRows.Count) | Access entries: $($AccessRows.Count) | Distinct users: $($UserAgg.Count)"
    return @{
        Workspaces  = $WsRows
        AccessRows  = $AccessRows
        UserRollup  = @($UserAgg.Values)
        SeenUpns    = $SeenUpns
    }
}

function Get-GraphToken {
    param([string]$TenantId)
    $Az = Get-Command az -ErrorAction SilentlyContinue
    if (-not $Az) { return $null }
    try {
        $Acct = (az account show 2>$null | ConvertFrom-Json)
        if (-not $Acct -or $Acct.tenantId -ne $TenantId) {
            Write-Dim "az login required for tenant $TenantId..."
            az login --tenant $TenantId --allow-no-subscriptions --only-show-errors | Out-Null
        }
        $Tok = (az account get-access-token --resource "https://graph.microsoft.com" --only-show-errors | ConvertFrom-Json).accessToken
        if (-not $Tok) { return $null }
        return "Bearer $Tok"
    } catch { return $null }
}

function Get-M365LicensedUsers {
    param([string]$TenantId, $SeenUpns)
    Write-Step "[4/12] Microsoft 365 Power BI / Fabric licenses"
    $Rows = @()
    if (-not $Script:DoGraph) { Add-Warn "M365 licenses" "Skipped by request."; return $Rows }

    $Token = Get-GraphToken -TenantId $TenantId
    if (-not $Token) {
        Add-Warn "M365 licenses" "No Graph token. Install Azure CLI and run 'az login' to enable this section."
        return $Rows
    }

    # skuId -> friendly name. Anything not listed is ignored.
    $SkuNames = @{
        "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
        "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Fabric (Free)"
        "b8a9ee8d-8a95-4f7c-82c8-6b43f5b6e67c" = "Power BI Premium Per User"
        "de376a03-6e0f-4d4c-b4cf-9b9a345b8a06" = "Power BI Premium Per User Add-On"
        "d05e6a75-3461-4c0f-9da8-31a9aac51b3d" = "Power BI Pro (GCC)"
        "3a6a908c-09c5-406a-8170-8ebb63c42882" = "Power BI Pro Dept"
        "7b26f5ab-a763-4c00-a1ac-f6c4b5506945" = "Power BI Premium P1"
        "c1d032e0-5619-4761-9b5c-75b6831e1711" = "Power BI Premium Per User Dept"
    }

    try {
        $Headers  = @{ Authorization = $Token; "Content-Type" = "application/json" }
        $Url      = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,userType,department,assignedLicenses&`$top=999"
        $AllUsers = @()
        do {
            $Resp     = Invoke-RestMethod -Uri $Url -Headers $Headers -Method Get -ErrorAction Stop
            $AllUsers += $Resp.value
            $Url      = $Resp.'@odata.nextLink'
        } while ($Url)
        Write-Dim "Graph users scanned: $($AllUsers.Count)"

        foreach ($U in $AllUsers) {
            $Upn = $U.userPrincipalName
            if (-not $Upn -or $Upn -notlike "*@*") { continue }

            $Matched = @()
            foreach ($L in @($U.assignedLicenses)) {
                if ($SkuNames.ContainsKey($L.skuId)) { $Matched += $SkuNames[$L.skuId] }
            }
            if ($Matched.Count -eq 0) { continue }

            $Rows += [PSCustomObject]@{
                UserPrincipalName   = $Upn
                DisplayName         = $U.displayName
                AccountEnabled      = [bool]$U.accountEnabled
                UserType            = $U.userType
                Department          = $U.department
                Licenses            = ($Matched | Sort-Object -Unique) -join "; "
                LicenseCount        = @($Matched | Sort-Object -Unique).Count
                InAnyWorkspaceAcl   = $SeenUpns.ContainsKey($Upn.ToLower())
            }
        }
        Write-Ok "$($Rows.Count) licensed users. In a workspace ACL: $(@($Rows | Where-Object { $_.InAnyWorkspaceAcl }).Count)"
    } catch {
        Add-Warn "M365 licenses" $_.Exception.Message
    }
    return $Rows
}

function Get-ContentInventory {
    param($Workspaces)
    Write-Step "[5/12] Reports, dashboards, dataflows, datasets"

    $Reports    = @()
    $Dashboards = @()
    $Dataflows  = @()
    $Datasets   = @()

    foreach ($Ws in $Workspaces) {
        foreach ($R in @($Ws.reports)) {
            $Reports += [PSCustomObject]@{
                WorkspaceId   = $Ws.id
                WorkspaceName = $Ws.name
                ReportId      = $R.id
                ReportName    = $R.name
                ReportType    = if ($R.reportType) { $R.reportType } else { "PowerBIReport" }
                DatasetId     = $R.datasetId
                CreatedBy     = $R.createdBy
                ModifiedBy    = $R.modifiedBy
                CreatedDate   = ConvertTo-IsoDate $R.createdDateTime
                ModifiedDate  = ConvertTo-IsoDate $R.modifiedDateTime
                WebUrl        = $R.webUrl
            }
        }
        foreach ($D in @($Ws.dashboards)) {
            $Dashboards += [PSCustomObject]@{
                WorkspaceId   = $Ws.id
                WorkspaceName = $Ws.name
                DashboardId   = $D.id
                DashboardName = $D.displayName
                IsReadOnly    = [bool]$D.isReadOnly
                WebUrl        = $D.webUrl
            }
        }
        foreach ($F in @($Ws.dataflows)) {
            $Dataflows += [PSCustomObject]@{
                WorkspaceId   = $Ws.id
                WorkspaceName = $Ws.name
                DataflowId    = $F.objectId
                DataflowName  = $F.name
                Description   = $F.description
                ConfiguredBy  = $F.configuredBy
                ModifiedBy    = $F.modifiedBy
                ModifiedDate  = ConvertTo-IsoDate $F.modifiedDateTime
            }
        }
        foreach ($S in @($Ws.datasets)) {
            $Storage = if ($S.targetStorageMode) { $S.targetStorageMode }
                       elseif ($S.defaultMode)   { $S.defaultMode }
                       else { "" }
            $Datasets += [PSCustomObject]@{
                WorkspaceId        = $Ws.id
                WorkspaceName      = $Ws.name
                DatasetId          = $S.id
                DatasetName        = $S.name
                ConfiguredBy       = $S.configuredBy
                StorageMode        = $Storage
                IsRefreshable      = [bool]$S.isRefreshable
                RequiresEffectiveIdentity = [bool]$S.isEffectiveIdentityRequired
                RequiresRoles      = [bool]$S.isEffectiveIdentityRolesRequired
                IsOnPremGatewayRequired   = [bool]$S.isOnPremGatewayRequired
                ContentProviderType = $S.contentProviderType
                CreatedDate        = ConvertTo-IsoDate $S.createdDate
                WebUrl             = $S.webUrl
            }
        }
    }

    Write-Ok "Reports: $($Reports.Count) | Dashboards: $($Dashboards.Count) | Dataflows: $($Dataflows.Count) | Datasets: $($Datasets.Count)"
    return @{
        Reports    = $Reports
        Dashboards = $Dashboards
        Dataflows  = $Dataflows
        Datasets   = $Datasets
    }
}

function Get-DatasetSources {
    param($Datasets)
    Write-Step "[6/12] Dataset data sources"
    $Rows = @()
    $i = 0
    $Failed = 0
    foreach ($D in $Datasets) {
        $i++
        if ($i % 200 -eq 0) { Write-Dim "dataset $i / $($Datasets.Count)" }
        $Resp = Invoke-PBIGet -Quiet -Url "https://api.powerbi.com/v1.0/myorg/admin/datasets/$($D.DatasetId)/datasources"
        if ($null -eq $Resp) { $Failed++; continue }
        foreach ($S in (Get-ValueArray $Resp)) {
            $CD = $S.connectionDetails
            $Parts = @()
            if ($CD) {
                foreach ($F in @('server','database','path','url','account','domain','kind','className')) {
                    if ($CD.PSObject.Properties[$F] -and $CD.$F) { $Parts += "$F=$($CD.$F)" }
                }
            }
            $Rows += [PSCustomObject]@{
                WorkspaceName  = $D.WorkspaceName
                DatasetId      = $D.DatasetId
                DatasetName    = $D.DatasetName
                DatasourceType = $S.datasourceType
                GatewayId      = $S.gatewayId
                DatasourceId   = $S.datasourceId
                ConnectionInfo = ($Parts -join " | ")
            }
        }
        Start-Sleep -Milliseconds ([math]::Max(80, $DelayMs / 2))
    }
    if ($Failed -gt 0) { Write-Dim "$Failed datasets did not return sources (deleted, or not readable)." }
    Write-Ok "$($Rows.Count) data source bindings."
    return $Rows
}

function Get-RefreshSchedules {
    param($Datasets)
    Write-Step "[7/12] Scheduled refresh configuration"
    $Rows = @()
    $Refreshable = @($Datasets | Where-Object { $_.IsRefreshable })
    $i = 0
    foreach ($D in $Refreshable) {
        $i++
        if ($i % 200 -eq 0) { Write-Dim "schedule $i / $($Refreshable.Count)" }
        $S = Invoke-PBIGet -Quiet -Url "https://api.powerbi.com/v1.0/myorg/admin/datasets/$($D.DatasetId)/refreshSchedule"
        if ($null -eq $S) { continue }
        $Times = @($S.times)
        $Rows += [PSCustomObject]@{
            WorkspaceName   = $D.WorkspaceName
            DatasetId       = $D.DatasetId
            DatasetName     = $D.DatasetName
            ScheduleEnabled = [bool]$S.enabled
            Frequency       = if ($S.frequency) { $S.frequency } else { "" }
            Days            = (@($S.days) -join ", ")
            Times           = ($Times -join ", ")
            RefreshesPerDay = $Times.Count
            TimeZone        = $S.localTimeZoneId
            NotifyOption    = $S.notifyOption
        }
        Start-Sleep -Milliseconds ([math]::Max(80, $DelayMs / 2))
    }
    Write-Ok "$($Rows.Count) schedules read ($(@($Rows | Where-Object { $_.ScheduleEnabled }).Count) enabled)."
    return $Rows
}

function Get-Refreshables {
    Write-Step "[8/12] Refreshable items (capacity refresh activity)"
    $Rows = @()
    try {
        $Raw = Invoke-PBIAdminList -Endpoint "capacities/refreshables" -QueryParams @{ '$expand' = 'capacity,group' }
        foreach ($R in $Raw) {
            $Last = $R.lastRefresh
            $Rows += [PSCustomObject]@{
                ItemId            = $R.id
                ItemName          = $R.name
                Kind              = $R.kind
                WorkspaceName     = if ($R.group) { $R.group.name } else { "" }
                CapacityName      = if ($R.capacity) { $R.capacity.displayName } else { "" }
                CapacitySku       = if ($R.capacity) { $R.capacity.sku } else { "" }
                RefreshCount      = $R.refreshCount
                RefreshFailures   = $R.refreshFailures
                AverageDurationSec = if ($null -ne $R.averageDuration) { [math]::Round([double]$R.averageDuration, 1) } else { $null }
                MedianDurationSec = if ($null -ne $R.medianDuration) { [math]::Round([double]$R.medianDuration, 1) } else { $null }
                RefreshesPerDay   = $R.refreshesPerDay
                LastRefreshStatus = if ($Last) { $Last.status } else { "" }
                LastRefreshStart  = if ($Last) { ConvertTo-IsoDate $Last.startTime } else { "" }
                LastRefreshEnd    = if ($Last) { ConvertTo-IsoDate $Last.endTime } else { "" }
                LastRefreshType   = if ($Last) { $Last.refreshType } else { "" }
                ConfiguredBy      = (@($R.configuredBy) -join "; ")
            }
        }
        Write-Ok "$($Rows.Count) refreshable items."
    } catch { Add-Warn "Refreshables" $_.Exception.Message }
    return $Rows
}

function Get-Gateways {
    Write-Step "[9/12] Gateways"
    $Rows = @()
    $SourceRows = @()
    try {
        $Raw = Get-ValueArray (Invoke-PBIGet -Url "https://api.powerbi.com/v1.0/myorg/admin/gateways")
        foreach ($G in $Raw) {
            $Type = switch ($G.type) {
                "Resource"       { "On-Premises (Standard)" }
                "VirtualNetwork" { "VNet" }
                "Personal"       { "Personal" }
                default          { $G.type }
            }
            $Srcs = Get-ValueArray (Invoke-PBIGet -Quiet -Url "https://api.powerbi.com/v1.0/myorg/admin/gateways/$($G.id)/datasources")
            foreach ($S in $Srcs) {
                $CD = $S.connectionDetails
                $SourceRows += [PSCustomObject]@{
                    GatewayId      = $G.id
                    GatewayName    = $G.name
                    DatasourceId   = $S.id
                    DatasourceName = $S.datasourceName
                    DatasourceType = $S.datasourceType
                    ConnectionDetails = if ($CD -is [string]) { $CD } else { ($CD | ConvertTo-Json -Compress -Depth 3) }
                }
            }
            Start-Sleep -Milliseconds $DelayMs

            $Rows += [PSCustomObject]@{
                GatewayId       = $G.id
                GatewayName     = $G.name
                GatewayType     = $Type
                GatewayVersion  = $G.gatewayVersion
                GatewayStatus   = $G.gatewayStatus
                GatewayMachine  = $G.gatewayMachine
                DataSourceCount = @($Srcs).Count
            }
        }
        Write-Ok "$($Rows.Count) gateways, $($SourceRows.Count) gateway data sources."
    } catch { Add-Warn "Gateways" $_.Exception.Message }
    return @{ Gateways = $Rows; GatewaySources = $SourceRows }
}

function Get-DeploymentPipelines {
    Write-Step "[10/12] Deployment pipelines"
    $Rows = @()
    try {
        $Raw = Get-ValueArray (Invoke-PBIGet -Url "https://api.powerbi.com/v1.0/myorg/admin/pipelines?`$expand=stages,users")
        foreach ($P in $Raw) {
            $Stages = @($P.stages)
            $Rows += [PSCustomObject]@{
                PipelineId   = $P.id
                PipelineName = $P.displayName
                Description  = $P.description
                StageCount   = $Stages.Count
                Stages       = (($Stages | Sort-Object order | ForEach-Object {
                                   $n = if ($_.workspaceName) { $_.workspaceName } else { "(empty)" }
                                   "$($_.order): $n" }) -join " -> ")
                AssignedWorkspaces = @($Stages | Where-Object { $_.workspaceId }).Count
                AdminCount   = if ($P.users) { @($P.users | Where-Object { $_.accessRight -eq 'Admin' }).Count } else { 0 }
            }
        }
        Write-Ok "$($Rows.Count) pipelines."
    } catch { Add-Warn "Deployment pipelines" $_.Exception.Message }
    return $Rows
}

function Get-PublishedApps {
    Write-Step "[11/12] Published apps"
    $Rows = @()
    try {
        $Raw = Get-ValueArray (Invoke-PBIGet -Url "https://api.powerbi.com/v1.0/myorg/admin/apps?`$top=5000")
        foreach ($A in $Raw) {
            $Rows += [PSCustomObject]@{
                AppId       = $A.id
                AppName     = $A.name
                Description = $A.description
                PublishedBy = $A.publishedBy
                WorkspaceId = $A.workspaceId
                LastUpdate  = ConvertTo-IsoDate $A.lastUpdate
            }
        }
        Write-Ok "$($Rows.Count) apps."
    } catch { Add-Warn "Published apps" $_.Exception.Message }
    return $Rows
}

function Get-TenantSettings {
    Write-Step "[12/12] Tenant settings (Fabric Admin API)"
    $Rows = @()
    try {
        $Token = (Get-PowerBIAccessToken -AsString -ErrorAction Stop)
        $Resp  = Invoke-RestMethod -Uri "https://api.fabric.microsoft.com/v1/admin/tenantsettings" `
                                   -Headers @{ Authorization = $Token } -Method Get -ErrorAction Stop
        foreach ($S in @($Resp.tenantSettings)) {
            $Groups = @($S.enabledSecurityGroups)
            $Rows += [PSCustomObject]@{
                SettingName   = $S.settingName
                Title         = $S.title
                Enabled       = [bool]$S.enabled
                CanSpecifySecurityGroups = [bool]$S.canSpecifySecurityGroups
                TenantSettingGroup = $S.tenantSettingGroup
                DelegatedFrom = $S.delegateToWorkspace
                SecurityGroups = (($Groups | ForEach-Object { $_.name }) -join "; ")
            }
        }
        Write-Ok "$($Rows.Count) tenant settings."
    } catch {
        Add-Warn "Tenant settings" "Not readable (needs Fabric Administrator). $($_.Exception.Message)"
    }
    return $Rows
}

function Invoke-ScannerScan {
    param($Workspaces)
    Write-Step "[+] Deep scan (endorsement + sensitivity labels)"
    $Rows = @()
    $WsIds = @($Workspaces | Where-Object { $_.type -ne 'PersonalGroup' } | ForEach-Object { $_.id })
    if ($WsIds.Count -eq 0) { return $Rows }

    $Base = "https://api.powerbi.com/v1.0/myorg/admin/workspaces"
    $Batches = [math]::Ceiling($WsIds.Count / 100)
    $B = 0
    for ($Start = 0; $Start -lt $WsIds.Count; $Start += 100) {
        $B++
        $Chunk = $WsIds[$Start..([math]::Min($Start + 99, $WsIds.Count - 1))]
        Write-Dim "scan batch $B / $Batches ($($Chunk.Count) workspaces)"
        try {
            $Body = @{ workspaces = $Chunk } | ConvertTo-Json -Depth 3
            $Url  = "$Base/getInfo?lineage=True&datasourceDetails=True&getArtifactUsers=True"
            $Scan = Invoke-PowerBIRestMethod -Url $Url -Method Post -Body $Body -ContentType "application/json" | ConvertFrom-Json
            $ScanId = $Scan.id

            $Status = ""
            $Tries  = 0
            while ($Status -ne "Succeeded" -and $Tries -lt 60) {
                Start-Sleep -Seconds 3
                $Tries++
                $St = Invoke-PBIGet -Quiet -Url "$Base/scanStatus/$ScanId"
                $Status = $St.status
                if ($Status -eq "Failed") { break }
            }
            if ($Status -ne "Succeeded") { Add-Warn "Deep scan" "Batch $B ended with status '$Status'."; continue }

            $Result = Invoke-PBIGet -Quiet -Url "$Base/scanResult/$ScanId"
            foreach ($W in @($Result.workspaces)) {
                foreach ($Kind in @('datasets','reports','dashboards','dataflows','datamarts','lakehouses','warehouses','notebooks')) {
                    if (-not $W.PSObject.Properties[$Kind]) { continue }
                    foreach ($It in @($W.$Kind)) {
                        $Endorse = if ($It.endorsementDetails) { $It.endorsementDetails.endorsement } else { "" }
                        $Rows += [PSCustomObject]@{
                            WorkspaceId   = $W.id
                            WorkspaceName = $W.name
                            ItemType      = $Kind.TrimEnd('s')
                            ItemId        = if ($It.id) { $It.id } else { $It.objectId }
                            ItemName      = if ($It.name) { $It.name } else { $It.displayName }
                            Endorsement   = $Endorse
                            CertifiedBy   = if ($It.endorsementDetails) { $It.endorsementDetails.certifiedBy } else { "" }
                            SensitivityLabelId = if ($It.sensitivityLabel) { $It.sensitivityLabel.labelId } else { "" }
                            Description   = $It.description
                            CreatedDate   = ConvertTo-IsoDate $It.createdDate
                            ModifiedDate  = ConvertTo-IsoDate $It.modifiedDateTime
                            ModifiedBy    = $It.modifiedBy
                        }
                    }
                }
            }
        } catch {
            Add-Warn "Deep scan" "Batch $B failed: $($_.Exception.Message)"
        }
    }
    Write-Ok "$($Rows.Count) items returned by scan."
    return $Rows
}


# ============================================================
# SUMMARY (counts only - no scoring)
# ============================================================
function Build-Summary {
    param($D)

    $Ws        = $D.Workspaces
    $Shared    = @($Ws | Where-Object { $_.Type -eq 'Workspace' -and $_.State -eq 'Active' })
    $Personal  = @($Ws | Where-Object { $_.Type -eq 'PersonalGroup' })
    $Refresh   = $D.Refreshables

    $ByCapacity = @($Shared | Group-Object CapacityName | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            CapacityName   = if ($_.Name) { $_.Name } else { "(no capacity - Pro / shared)" }
            WorkspaceCount = $_.Count
            DatasetCount   = ($_.Group | Measure-Object DatasetCount -Sum).Sum
            ReportCount    = ($_.Group | Measure-Object ReportCount -Sum).Sum
        }
    })

    $ByLicenseMode = @($Shared | Group-Object LicenseMode | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            LicenseMode    = if ($_.Name) { $_.Name } else { "(not reported)" }
            WorkspaceCount = $_.Count
        }
    })

    $ByStorageMode = @($D.Datasets | Group-Object StorageMode | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            StorageMode  = if ($_.Name) { $_.Name } else { "(not reported)" }
            DatasetCount = $_.Count
        }
    })

    $BySourceType = @($D.DatasetSources | Group-Object DatasourceType | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{
            DatasourceType = if ($_.Name) { $_.Name } else { "(not reported)" }
            BindingCount   = $_.Count
            DatasetCount   = @($_.Group | Select-Object -ExpandProperty DatasetId -Unique).Count
        }
    })

    $ByRole = @($D.AccessEntries | Group-Object Role | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{ Role = $_.Name; EntryCount = $_.Count }
    })

    $ByPrincipalType = @($D.AccessEntries | Group-Object PrincipalType | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{ PrincipalType = $_.Name; EntryCount = $_.Count }
    })

    $LastStatus = @($Refresh | Where-Object { $_.LastRefreshStatus } | Group-Object LastRefreshStatus | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{ LastRefreshStatus = $_.Name; ItemCount = $_.Count }
    })

    $Counts = [ordered]@{
        Capacities            = @($D.Capacities).Count
        WorkspacesActiveShared = $Shared.Count
        WorkspacesPersonal    = $Personal.Count
        WorkspacesOnCapacity  = @($Shared | Where-Object { $_.CapacityId }).Count
        WorkspacesNoContent   = @($Shared | Where-Object { $_.HasNoContent }).Count
        WorkspacesNoAdmin     = @($Shared | Where-Object { -not $_.HasWorkspaceAdmin }).Count
        AccessEntries         = @($D.AccessEntries).Count
        DistinctUsers         = @($D.UserRollup).Count
        GuestUsers            = @($D.UserRollup | Where-Object { $_.IsGuest }).Count
        M365LicensedUsers     = @($D.M365Users).Count
        Datasets              = @($D.Datasets).Count
        DatasetsRefreshable   = @($D.Datasets | Where-Object { $_.IsRefreshable }).Count
        DatasetsWithRLS       = @($D.Datasets | Where-Object { $_.RequiresEffectiveIdentity }).Count
        DataSourceBindings    = @($D.DatasetSources).Count
        RefreshSchedules      = @($D.RefreshSchedules).Count
        RefreshSchedulesEnabled = @($D.RefreshSchedules | Where-Object { $_.ScheduleEnabled }).Count
        RefreshableItems      = @($Refresh).Count
        Reports               = @($D.Reports | Where-Object { $_.ReportType -ne 'PaginatedReport' }).Count
        PaginatedReports      = @($D.Reports | Where-Object { $_.ReportType -eq 'PaginatedReport' }).Count
        Dashboards            = @($D.Dashboards).Count
        Dataflows             = @($D.Dataflows).Count
        Gateways              = @($D.Gateways).Count
        GatewayDataSources    = @($D.GatewaySources).Count
        DeploymentPipelines   = @($D.Pipelines).Count
        PublishedApps         = @($D.Apps).Count
        TenantSettingsRead    = @($D.TenantSettings).Count
        ScannedItems          = @($D.ScanItems).Count
    }

    return [PSCustomObject]@{
        Counts          = $Counts
        ByCapacity      = $ByCapacity
        ByLicenseMode   = $ByLicenseMode
        ByStorageMode   = $ByStorageMode
        BySourceType    = $BySourceType
        ByRole          = $ByRole
        ByPrincipalType = $ByPrincipalType
        ByLastRefreshStatus = $LastStatus
    }
}

# ============================================================
# EXPORT
# ============================================================
function Export-JsonReport {
    param($Payload, [string]$Path)
    $Payload | ConvertTo-Json -Depth 12 | Out-File -FilePath $Path -Encoding UTF8
    Write-Ok "JSON : $Path"
}

function Export-CsvFiles {
    param($D, [string]$Folder)
    if (-not (Test-Path $Folder)) { New-Item -ItemType Directory -Path $Folder | Out-Null }
    $Map = @{
        "capacities"        = $D.Capacities
        "workspaces"        = $D.Workspaces
        "workspace_access"  = $D.AccessEntries
        "user_rollup"       = $D.UserRollup
        "m365_licenses"     = $D.M365Users
        "datasets"          = $D.Datasets
        "dataset_sources"   = $D.DatasetSources
        "refresh_schedules" = $D.RefreshSchedules
        "refreshable_items" = $D.Refreshables
        "reports"           = $D.Reports
        "dashboards"        = $D.Dashboards
        "dataflows"         = $D.Dataflows
        "gateways"          = $D.Gateways
        "gateway_sources"   = $D.GatewaySources
        "pipelines"         = $D.Pipelines
        "apps"              = $D.Apps
        "tenant_settings"   = $D.TenantSettings
        "scanned_items"     = $D.ScanItems
    }
    $N = 0
    foreach ($K in $Map.Keys) {
        $Rows = @($Map[$K])
        if ($Rows.Count -eq 0) { continue }
        $F = Join-Path $Folder "$K.csv"
        $Rows | Export-Csv -Path $F -NoTypeInformation -Encoding UTF8
        $N++
    }
    Write-Ok "CSV  : $N files in $Folder"
}

function ConvertTo-JsSafeJson {
    param($Data)
    $Json = $Data | ConvertTo-Json -Depth 12 -Compress
    if (-not $Json) { $Json = "[]" }
    # Guard against the JSON string terminating the inline <script> block.
    return $Json.Replace("</", "<\/")
}

function Export-HtmlReport {
    param($Payload, [string]$Path)

    $D = $Payload.Data
    $S = $Payload.Summary

    # Each tab: key, label, rows. Order matters - it drives the nav.
    $Tabs = [ordered]@{
        "capacities"        = @{ Label = "Capacities";        Rows = $D.Capacities }
        "workspaces"        = @{ Label = "Workspaces";        Rows = $D.Workspaces }
        "workspace_access"  = @{ Label = "Workspace Access";  Rows = $D.AccessEntries }
        "user_rollup"       = @{ Label = "Users";             Rows = $D.UserRollup }
        "m365_licenses"     = @{ Label = "M365 Licenses";     Rows = $D.M365Users }
        "datasets"          = @{ Label = "Semantic Models";   Rows = $D.Datasets }
        "dataset_sources"   = @{ Label = "Data Sources";      Rows = $D.DatasetSources }
        "refresh_schedules" = @{ Label = "Refresh Schedules"; Rows = $D.RefreshSchedules }
        "refreshable_items" = @{ Label = "Refresh Activity";  Rows = $D.Refreshables }
        "reports"           = @{ Label = "Reports";           Rows = $D.Reports }
        "dashboards"        = @{ Label = "Dashboards";        Rows = $D.Dashboards }
        "dataflows"         = @{ Label = "Dataflows";         Rows = $D.Dataflows }
        "gateways"          = @{ Label = "Gateways";          Rows = $D.Gateways }
        "gateway_sources"   = @{ Label = "Gateway Sources";   Rows = $D.GatewaySources }
        "pipelines"         = @{ Label = "Pipelines";         Rows = $D.Pipelines }
        "apps"              = @{ Label = "Apps";              Rows = $D.Apps }
        "tenant_settings"   = @{ Label = "Tenant Settings";   Rows = $D.TenantSettings }
        "scanned_items"     = @{ Label = "Scanned Items";     Rows = $D.ScanItems }
    }

    $TabPayload = [ordered]@{}
    foreach ($K in $Tabs.Keys) {
        $Rows = @($Tabs[$K].Rows)
        $Cols = @()
        if ($Rows.Count -gt 0) {
            $Cols = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        }
        $TabPayload[$K] = [ordered]@{
            label = $Tabs[$K].Label
            count = $Rows.Count
            cols  = $Cols
            rows  = $Rows
        }
    }

    $DataJson    = ConvertTo-JsSafeJson $TabPayload
    $SummaryJson = ConvertTo-JsSafeJson $S
    $MetaJson    = ConvertTo-JsSafeJson $Payload.Meta
    $WarnJson    = ConvertTo-JsSafeJson @($Payload.Warnings)

    $Html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Power BI Environment Inventory</title>
<style>
  :root {
    --bg:#f5f6f8; --panel:#ffffff; --ink:#1b1b1f; --muted:#5f6368;
    --line:#e3e5e8; --accent:#0b6a9e; --accent-soft:#e8f2f8;
  }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--ink);
         font:14px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif; }
  header { background:var(--panel); border-bottom:1px solid var(--line); padding:20px 28px; }
  h1 { margin:0 0 4px; font-size:20px; font-weight:600; }
  .sub { color:var(--muted); font-size:13px; }
  .wrap { padding:20px 28px 60px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fill,minmax(180px,1fr)); gap:10px; margin-bottom:22px; }
  .card { background:var(--panel); border:1px solid var(--line); border-radius:6px; padding:12px 14px; }
  .card .n { font-size:22px; font-weight:600; }
  .card .l { font-size:11px; color:var(--muted); text-transform:uppercase; letter-spacing:.4px; margin-top:2px; }
  .panel { background:var(--panel); border:1px solid var(--line); border-radius:6px; padding:16px 18px; margin-bottom:22px; }
  .panel h2 { margin:0 0 12px; font-size:14px; font-weight:600; }
  .grid2 { display:grid; grid-template-columns:repeat(auto-fit,minmax(300px,1fr)); gap:18px; }
  nav { display:flex; flex-wrap:wrap; gap:4px; margin-bottom:14px; }
  nav button { background:var(--panel); border:1px solid var(--line); border-radius:4px;
               padding:6px 11px; cursor:pointer; font-size:13px; color:var(--ink); }
  nav button:hover { background:var(--accent-soft); }
  nav button.on { background:var(--accent); border-color:var(--accent); color:#fff; }
  nav button .c { opacity:.65; font-size:11px; margin-left:5px; }
  .toolbar { display:flex; gap:8px; align-items:center; margin-bottom:10px; flex-wrap:wrap; }
  input[type=search] { flex:1; min-width:220px; padding:7px 10px; border:1px solid var(--line);
                       border-radius:4px; font-size:13px; }
  .btn { background:var(--panel); border:1px solid var(--line); border-radius:4px;
         padding:7px 12px; cursor:pointer; font-size:13px; }
  .btn:hover { background:var(--accent-soft); }
  .shown { color:var(--muted); font-size:12px; }
  .tablebox { overflow:auto; max-height:70vh; border:1px solid var(--line); border-radius:6px; background:var(--panel); }
  table { border-collapse:collapse; width:100%; font-size:12.5px; }
  th { position:sticky; top:0; background:#eef0f3; text-align:left; padding:8px 10px;
       border-bottom:1px solid var(--line); white-space:nowrap; cursor:pointer; font-weight:600; }
  th:hover { background:#e2e6ea; }
  td { padding:6px 10px; border-bottom:1px solid #f0f1f3; vertical-align:top;
       max-width:420px; overflow-wrap:anywhere; }
  tr:hover td { background:#fafbfc; }
  .empty { padding:26px; text-align:center; color:var(--muted); }
  .mini { width:100%; border-collapse:collapse; font-size:12.5px; }
  .mini td, .mini th { padding:5px 8px; border-bottom:1px solid #f0f1f3; text-align:left; }
  .mini th { background:transparent; position:static; cursor:default; }
  .mini td:last-child { text-align:right; font-variant-numeric:tabular-nums; }
  .warn { background:#fff8e6; border:1px solid #f0dda0; border-radius:6px; padding:12px 16px; margin-bottom:22px; }
  .warn h2 { margin:0 0 8px; font-size:13px; }
  .warn li { font-size:12.5px; color:#6b5a1e; }
  footer { color:var(--muted); font-size:12px; padding:0 28px 40px; }
  code { background:#eef0f3; padding:1px 5px; border-radius:3px; font-size:12px; }
</style>
</head>
<body>
<header>
  <h1>Power BI Environment Inventory</h1>
  <div class="sub" id="meta"></div>
</header>
<div class="wrap">
  <div id="warnings"></div>
  <div class="cards" id="cards"></div>

  <div class="grid2">
    <div class="panel"><h2>Workspaces by capacity</h2><div id="t_cap"></div></div>
    <div class="panel"><h2>Workspaces by license mode</h2><div id="t_lic"></div></div>
    <div class="panel"><h2>Semantic models by storage mode</h2><div id="t_sto"></div></div>
    <div class="panel"><h2>Data source types</h2><div id="t_src"></div></div>
    <div class="panel"><h2>Workspace access by role</h2><div id="t_role"></div></div>
    <div class="panel"><h2>Access entries by principal type</h2><div id="t_ptype"></div></div>
    <div class="panel"><h2>Last refresh outcome</h2><div id="t_ref"></div></div>
  </div>

  <nav id="nav"></nav>
  <div class="toolbar">
    <input type="search" id="q" placeholder="Filter rows in this table...">
    <button class="btn" id="csv">Download CSV</button>
    <span class="shown" id="shown"></span>
  </div>
  <div class="tablebox"><div id="table"></div></div>
</div>
<footer>
  <b>Read-only report.</b> The script that produced this file only reads inventory data from the
  Power BI Admin REST API, Microsoft Graph, and the Fabric Admin API. It created, modified and
  deleted nothing in the tenant. Run it with <code>-ListApiCalls</code> to list every endpoint it
  can call and the HTTP verb for each.
  <br><br>
  This report is a factual census: it contains no scoring, ranking, or recommendations. Every value
  shown is returned by an API or is a direct count of API values.
</footer>

<script>
const DATA = $DataJson;
const SUM  = $SummaryJson;
const META = $MetaJson;
const WARN = $WarnJson;

const el = id => document.getElementById(id);
const esc = v => (v===null||v===undefined) ? "" : String(v)
  .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
// PowerShell's ConvertTo-Json unrolls a single-element array into a bare object.
// Normalize everything we iterate so a one-row collection still renders.
const arr = v => Array.isArray(v) ? v : (v === null || v === undefined ? [] : [v]);
for (const k of Object.keys(DATA)) { DATA[k].rows = arr(DATA[k].rows); DATA[k].cols = arr(DATA[k].cols); }

el("meta").textContent =
  "Tenant " + META.TenantId + "  |  generated " + META.GeneratedAt +
  "  |  by " + META.GeneratedBy + "  |  deep scan: " + (META.DeepScan ? "yes" : "no");

if (arr(WARN).length) {
  el("warnings").innerHTML = '<div class="warn"><h2>Sections that could not be collected</h2><ul>' +
    arr(WARN).map(w => "<li><b>" + esc(w.Area) + "</b>: " + esc(w.Message) + "</li>").join("") +
    "</ul></div>";
}

const LABELS = {
  Capacities:"Capacities", WorkspacesActiveShared:"Active workspaces", WorkspacesPersonal:"Personal workspaces",
  WorkspacesOnCapacity:"On dedicated capacity", WorkspacesNoContent:"Workspaces with no items",
  WorkspacesNoAdmin:"Workspaces with no admin", AccessEntries:"Access entries", DistinctUsers:"Distinct users",
  GuestUsers:"Guest users", M365LicensedUsers:"M365 licensed users", Datasets:"Semantic models",
  DatasetsRefreshable:"Refreshable models", DatasetsWithRLS:"Models requiring identity",
  DataSourceBindings:"Data source bindings", RefreshSchedules:"Refresh schedules",
  RefreshSchedulesEnabled:"Schedules enabled", RefreshableItems:"Items with refresh history",
  Reports:"Reports", PaginatedReports:"Paginated reports", Dashboards:"Dashboards", Dataflows:"Dataflows",
  Gateways:"Gateways", GatewayDataSources:"Gateway data sources", DeploymentPipelines:"Deployment pipelines",
  PublishedApps:"Published apps", TenantSettingsRead:"Tenant settings read", ScannedItems:"Scanned items"
};
el("cards").innerHTML = Object.keys(SUM.Counts).map(k =>
  '<div class="card"><div class="n">' + esc(SUM.Counts[k]) + '</div><div class="l">' +
  esc(LABELS[k] || k) + "</div></div>").join("");

function miniTable(target, rows) {
  const r = arr(rows);
  if (!r.length) { el(target).innerHTML = '<div class="empty">No data.</div>'; return; }
  const cols = Object.keys(r[0]);
  el(target).innerHTML = '<table class="mini"><thead><tr>' +
    cols.map(c => "<th>" + esc(c) + "</th>").join("") + "</tr></thead><tbody>" +
    r.map(x => "<tr>" + cols.map(c => "<td>" + esc(x[c]) + "</td>").join("") + "</tr>").join("") +
    "</tbody></table>";
}
miniTable("t_cap",   SUM.ByCapacity);
miniTable("t_lic",   SUM.ByLicenseMode);
miniTable("t_sto",   SUM.ByStorageMode);
miniTable("t_src",   SUM.BySourceType);
miniTable("t_role",  SUM.ByRole);
miniTable("t_ptype", SUM.ByPrincipalType);
miniTable("t_ref",   SUM.ByLastRefreshStatus);

const keys = Object.keys(DATA);
let active = keys.find(k => DATA[k].count > 0) || keys[0];
let sortCol = null, sortDir = 1;

el("nav").innerHTML = keys.map(k =>
  '<button data-k="' + k + '">' + esc(DATA[k].label) +
  '<span class="c">' + DATA[k].count + "</span></button>").join("");
el("nav").addEventListener("click", e => {
  const b = e.target.closest("button");
  if (!b) return;
  active = b.dataset.k; sortCol = null; el("q").value = ""; render();
});

function visibleRows() {
  const t = DATA[active];
  const q = el("q").value.trim().toLowerCase();
  let rows = t.rows.slice();
  if (q) rows = rows.filter(r => t.cols.some(c =>
    String(r[c] === null || r[c] === undefined ? "" : r[c]).toLowerCase().includes(q)));
  if (sortCol) {
    rows.sort((a, b) => {
      const x = a[sortCol], y = b[sortCol];
      const nx = parseFloat(x), ny = parseFloat(y);
      if (!isNaN(nx) && !isNaN(ny) && String(x).trim() !== "" && String(y).trim() !== "")
        return (nx - ny) * sortDir;
      return String(x === null || x === undefined ? "" : x)
        .localeCompare(String(y === null || y === undefined ? "" : y)) * sortDir;
    });
  }
  return rows;
}

function render() {
  document.querySelectorAll("#nav button").forEach(b =>
    b.classList.toggle("on", b.dataset.k === active));
  const t = DATA[active];
  if (!t.count) {
    el("table").innerHTML = '<div class="empty">No rows collected for ' + esc(t.label) + ".</div>";
    el("shown").textContent = ""; return;
  }
  const rows = visibleRows();
  el("table").innerHTML = "<table><thead><tr>" +
    t.cols.map(c => "<th data-c='" + esc(c) + "'>" + esc(c) +
      (sortCol === c ? (sortDir > 0 ? " \u25B2" : " \u25BC") : "") + "</th>").join("") +
    "</tr></thead><tbody>" +
    rows.map(r => "<tr>" + t.cols.map(c => "<td>" + esc(r[c]) + "</td>").join("") + "</tr>").join("") +
    "</tbody></table>";
  el("shown").textContent = rows.length + " of " + t.count + " rows";
  document.querySelectorAll("#table th").forEach(th => th.onclick = () => {
    const c = th.dataset.c;
    if (sortCol === c) sortDir = -sortDir; else { sortCol = c; sortDir = 1; }
    render();
  });
}

el("q").addEventListener("input", render);
el("csv").addEventListener("click", () => {
  const t = DATA[active], rows = visibleRows();
  const q = v => '"' + String(v === null || v === undefined ? "" : v).replace(/"/g, '""') + '"';
  const csv = [t.cols.map(q).join(",")]
    .concat(rows.map(r => t.cols.map(c => q(r[c])).join(","))).join("\r\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" }));
  a.download = "pbi_" + active + ".csv";
  a.click();
});

render();
</script>
</body>
</html>
"@

    $Html | Out-File -FilePath $Path -Encoding UTF8
    Write-Ok "HTML : $Path"
}

# ============================================================
# MAIN
# ============================================================
try {
    Read-Config
    Connect-PBIService -TenantId $Script:TenantId

    $Started = Get-Date

    $Capacities   = Get-Capacities
    $RawWs        = Get-AllWorkspacesRaw
    $WsInv        = Build-WorkspaceInventory -Workspaces $RawWs -CapacityRows $Capacities
    $M365Users    = Get-M365LicensedUsers -TenantId $Script:TenantId -SeenUpns $WsInv.SeenUpns
    $Content      = Get-ContentInventory -Workspaces $RawWs
    $DsSources    = Get-DatasetSources -Datasets $Content.Datasets
    $Schedules    = Get-RefreshSchedules -Datasets $Content.Datasets
    $Refreshables = Get-Refreshables
    $Gw           = Get-Gateways
    $Pipelines    = Get-DeploymentPipelines
    $Apps         = Get-PublishedApps
    $Settings     = Get-TenantSettings
    $ScanItems    = if ($Script:DoDeepScan) { Invoke-ScannerScan -Workspaces $RawWs } else { @() }

    $Data = [ordered]@{
        Capacities       = $Capacities
        Workspaces       = $WsInv.Workspaces
        AccessEntries    = $WsInv.AccessRows
        UserRollup       = $WsInv.UserRollup
        M365Users        = $M365Users
        Datasets         = $Content.Datasets
        DatasetSources   = $DsSources
        RefreshSchedules = $Schedules
        Refreshables     = $Refreshables
        Reports          = $Content.Reports
        Dashboards       = $Content.Dashboards
        Dataflows        = $Content.Dataflows
        Gateways         = $Gw.Gateways
        GatewaySources   = $Gw.GatewaySources
        Pipelines        = $Pipelines
        Apps             = $Apps
        TenantSettings   = $Settings
        ScanItems        = $ScanItems
    }

    Write-Host ""
    Write-Step "Building report"
    $Summary = Build-Summary -D $Data

    $Payload = [ordered]@{
        Meta = [ordered]@{
            TenantId      = $Script:TenantId
            GeneratedAt   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            GeneratedBy   = $env:USERNAME
            DurationMin   = [math]::Round(((Get-Date) - $Started).TotalMinutes, 1)
            DeepScan      = [bool]$Script:DoDeepScan
            GraphIncluded = [bool]$Script:DoGraph
            ScriptVersion = "1.0.0"
        }
        Summary  = $Summary
        Data     = $Data
        Warnings = @($Script:Warnings)
    }

    $JsonPath = Join-Path $OutputFolder "PBI_Environment_Inventory.json"
    $HtmlPath = Join-Path $OutputFolder "PBI_Environment_Inventory.html"
    Export-JsonReport -Payload $Payload -Path $JsonPath
    Export-HtmlReport -Payload $Payload -Path $HtmlPath
    if ($Script:DoCsv) { Export-CsvFiles -D $Data -Folder (Join-Path $OutputFolder "csv") }

    Write-Host ""
    Write-Host "=========================================================" -ForegroundColor Cyan
    Write-Host "  INVENTORY COMPLETE" -ForegroundColor Cyan
    Write-Host "=========================================================" -ForegroundColor Cyan
    foreach ($K in $Summary.Counts.Keys) {
        Write-Host ("  {0,-26} {1}" -f $K, $Summary.Counts[$K]) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "  Elapsed: $($Payload.Meta.DurationMin) min" -ForegroundColor DarkGray
    if ($Script:Warnings.Count -gt 0) {
        Write-Host "  Sections skipped: $($Script:Warnings.Count) (listed at the top of the HTML report)" -ForegroundColor DarkYellow
    }
    Write-Host ""
    Write-Host "  Open: $HtmlPath" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}

