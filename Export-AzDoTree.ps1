<#
================================================================================
 Export-AzDoTree.ps1
 Download een deel van de Azure DevOps work-item tree en exporteer naar Excel.
================================================================================

 DESIGN DECISIONS
 ----------------
 1. Auth via Personal Access Token (PAT), header = "Basic :<pat>" (b64).
    Why: standaard voor AzDO REST, werkt in CI/CD en lokaal, geen interactive
    login nodig. PAT uit env var $env:AZDO_PAT of parameter -Pat.
 2. Roots bepalen via WIQL (-Wiql / -WiqlFile) of expliciete -RootIds.
    Why: elk team heeft andere filters (area path, iteratie, type, state);
    WIQL houdt filterlogica buiten het script.
 3. Children ophalen via /workitems/{id}?$expand=Relations (System.LinkTypes.
    Hierarchy-Forward). Why: WIQL "WorkItemLinks" geeft wel relaties maar is
    duurder bij deep trees met weinig roots; per-root expansie is simpeler en
    schaalt tot MaxDepth.
 4. Velden bulk ophalen via POST /wit/workitemsbatch (max 200 ids/call).
    Why: 1 call per 200 items i.p.v. 200 calls; essentieel voor >1000 items.
 5. Export via ImportExcel-module (geen Excel COM nodig).
    Why: headless/server-ready, auto-size + freeze + autofilter in 1 call.
 6. Tree-indent via Title-prefix ("  " * Level) + aparte Level-kolom.
    Why: Excel outline-grouping werkt matig vanuit PS; indented title + Level
    is zowel visueel als machine-leesbaar.
 7. Retry met Retry-After header respect.
    Why: AzDO throttlet agressief bij >200 req/min; 429 negeren = half-leeg
    resultaat zonder duidelijke error.
 8. Default field-whitelist: id, type, title, state, assignedTo, iterationPath,
    areaPath, storyPoints, tags, parent, changedDate.
    Why: dekt 95% van rapportage-use-cases; -Fields override voor rest.

 ARCHITECTURE
 ------------
   +-------------------------------------------------+
   |  Export-AzDoTree.ps1                            |
   |                                                 |
   |  [Params] --> [Auth header] --> [HTTP client]   |
   |                                       |         |
   |                                       v         |
   |  [WIQL / RootIds] --> [Roots: ids + depth]      |
   |                                       |         |
   |                                       v         |
   |  [Recursive expand (Relations)] --> [Full tree] |
   |                                       |         |
   |                                       v         |
   |  [Bulk /workitemsbatch (200/chunk)] --> [Fields]|
   |                                       |         |
   |                                       v         |
   |  [Flatten + indent + sort] --> [Row objects]    |
   |                                       |         |
   |                                       v         |
   |  [Export-Excel (freeze/filter/autosize)]        |
   +-------------------------------------------------+
                             |
                             v
                       tree.xlsx

 DATAFLOW
 --------
   user ---1. params+PAT--> script
   script ---2. POST /wiql--> AzDO ---workItems[]---> script
   script ---3. GET /workitems/{id}?$expand=Relations--> AzDO
                   (loop over levels, dedup ids, until MaxDepth)
   script ---4. POST /workitemsbatch (chunks)--> AzDO ---fields[]--> script
   script ---5. join id->fields, build rows (Level,Id,Type,Title,...)
   script ---6. Export-Excel---> tree.xlsx on disk

 DEPENDENCIES
 ------------
   * PowerShell 5.1+ (7.x aanbevolen i.v.m. HttpClient perf).
   * Module ImportExcel  (Install-Module ImportExcel -Scope CurrentUser).
   * Azure DevOps PAT met scope "Work Items (Read)".
   * Netwerk naar dev.azure.com (of on-prem TFS/AzDO Server host).
   * Env var AZDO_PAT  (of parameter -Pat).
================================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]   $Organization,
    [Parameter(Mandatory)] [string]   $Project,

    [string]                          $Wiql,
    [string]                          $WiqlFile,
    [int[]]                           $RootIds,

    [string[]]                        $Fields = @(
        'System.Id','System.WorkItemType','System.Title','System.State',
        'System.AssignedTo','System.IterationPath','System.AreaPath',
        'System.Tags','System.Parent','System.ChangedDate',
        'Microsoft.VSTS.Scheduling.StoryPoints'
    ),

    [int]                             $MaxDepth = 5,
    [string]                          $OutFile  = "./tree.xlsx",
    [string]                          $SheetName = "Tree",
    [string]                          $Pat = $env:AZDO_PAT,
    [string]                          $BaseUrl = "https://dev.azure.com",
    [string]                          $ApiVersion = "7.1",
    [switch]                          $DryRun
)

$ErrorActionPreference = 'Stop'

# --- preflight -----------------------------------------------------------
if (-not $Pat) { throw "PAT ontbreekt. Zet \$env:AZDO_PAT of geef -Pat mee." }
if (-not (Get-Module -ListAvailable ImportExcel)) {
    throw "Module ImportExcel ontbreekt. Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel -ErrorAction Stop

if (-not $Wiql -and -not $WiqlFile -and -not $RootIds) {
    throw "Geef -Wiql, -WiqlFile of -RootIds mee om de tree-roots te bepalen."
}
if ($WiqlFile) { $Wiql = Get-Content -Raw -Path $WiqlFile }

$authHeader = @{
    Authorization = "Basic " + [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes(":$Pat"))
    'Content-Type' = 'application/json'
}
$projUrl = "$BaseUrl/$Organization/$Project/_apis"

# --- HTTP helper met 429/5xx retry ---------------------------------------
function Invoke-AzDo {
    param(
        [Parameter(Mandatory)] [string] $Method,
        [Parameter(Mandatory)] [string] $Uri,
        $Body
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $params = @{
                Method = $Method; Uri = $Uri; Headers = $authHeader
                UseBasicParsing = $true
            }
            if ($Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10) }
            return Invoke-RestMethod @params
        } catch {
            $resp = $_.Exception.Response
            $code = if ($resp) { [int]$resp.StatusCode } else { 0 }
            if ($code -in 429,500,502,503,504 -and $attempt -lt 5) {
                $wait = 2 * $attempt
                if ($resp -and $resp.Headers['Retry-After']) {
                    $wait = [int]$resp.Headers['Retry-After']
                }
                Write-Warning "AzDO $code; retry $attempt in ${wait}s"
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

# --- 1. roots -------------------------------------------------------------
$rootSet = [System.Collections.Generic.HashSet[int]]::new()

if ($RootIds) { foreach ($id in $RootIds) { [void]$rootSet.Add($id) } }

if ($Wiql) {
    Write-Host "WIQL query uitvoeren..."
    $wiqlResp = Invoke-AzDo -Method POST `
        -Uri "$projUrl/wit/wiql?api-version=$ApiVersion" `
        -Body @{ query = $Wiql }

    if ($wiqlResp.workItems) {
        foreach ($w in $wiqlResp.workItems) { [void]$rootSet.Add([int]$w.id) }
    } elseif ($wiqlResp.workItemRelations) {
        foreach ($r in $wiqlResp.workItemRelations) {
            if ($r.target) { [void]$rootSet.Add([int]$r.target.id) }
        }
    }
}

Write-Host "Roots gevonden: $($rootSet.Count)"
if ($DryRun) {
    $rootSet | Sort-Object | Select-Object -First 50 |
        ForEach-Object { "  $_" } | Write-Host
    Write-Host "(dry-run; stop)"; return
}
if ($rootSet.Count -eq 0) { throw "Geen roots om te exporteren." }

# --- 2. recursieve child-expansie ----------------------------------------
#    id -> @{ level=<int>; parent=<int?> }
$node       = @{}
$parentOf   = @{}
foreach ($id in $rootSet) { $node[$id] = @{ level = 0; parent = $null } }

$frontier = @($rootSet)
for ($depth = 0; $depth -lt $MaxDepth; $depth++) {
    if (-not $frontier) { break }
    Write-Host ("Expand level {0} ({1} items)" -f $depth, $frontier.Count)

    $next = New-Object System.Collections.Generic.List[int]
    foreach ($pid in $frontier) {
        $wi = Invoke-AzDo -Method GET `
            -Uri "$projUrl/wit/workitems/$pid`?`$expand=Relations&api-version=$ApiVersion"

        if (-not $wi.relations) { continue }
        foreach ($rel in $wi.relations) {
            if ($rel.rel -ne 'System.LinkTypes.Hierarchy-Forward') { continue }
            $cid = [int]($rel.url -split '/')[-1]
            if ($node.ContainsKey($cid)) { continue }
            $node[$cid] = @{ level = $depth + 1; parent = $pid }
            $parentOf[$cid] = $pid
            $next.Add($cid)
        }
    }
    $frontier = $next.ToArray()
}

$allIds = $node.Keys
Write-Host "Totaal items in tree: $($allIds.Count)"

# --- 3. bulk fields --------------------------------------------------------
$fieldsById = @{}
$chunks = [math]::Ceiling($allIds.Count / 200.0)
for ($i = 0; $i -lt $chunks; $i++) {
    $slice = $allIds | Select-Object -Skip ($i*200) -First 200
    $body = @{ ids = @($slice); fields = $Fields }
    Write-Host ("Batch {0}/{1} ({2} ids)" -f ($i+1), $chunks, $slice.Count)
    $resp = Invoke-AzDo -Method POST `
        -Uri "$projUrl/wit/workitemsbatch?api-version=$ApiVersion" `
        -Body $body
    foreach ($w in $resp.value) { $fieldsById[[int]$w.id] = $w.fields }
}

# --- 4. flatten -----------------------------------------------------------
function Get-AssignedToName($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [string]) { return $v }
    if ($v.displayName) { return $v.displayName }
    return "$v"
}

$rows = foreach ($id in $allIds) {
    $f = $fieldsById[$id]; if (-not $f) { continue }
    $lvl = $node[$id].level
    $title = ("  " * $lvl) + ($f.'System.Title')
    [pscustomobject]@{
        Level          = $lvl
        Id             = $id
        ParentId       = $node[$id].parent
        Type           = $f.'System.WorkItemType'
        Title          = $title
        State          = $f.'System.State'
        AssignedTo     = Get-AssignedToName $f.'System.AssignedTo'
        StoryPoints    = $f.'Microsoft.VSTS.Scheduling.StoryPoints'
        IterationPath  = $f.'System.IterationPath'
        AreaPath       = $f.'System.AreaPath'
        Tags           = $f.'System.Tags'
        ChangedDate    = $f.'System.ChangedDate'
        Url            = "$BaseUrl/$Organization/$Project/_workitems/edit/$id"
    }
}

# sort: parent-chain preorder zodat tree-volgorde klopt in Excel
function Get-SortKey($row, $byParent) {
    $chain = @()
    $cur = $row
    while ($cur) {
        $chain = ,([int]$cur.Id) + $chain
        if ($null -eq $cur.ParentId) { break }
        $cur = $byParent[$cur.ParentId]
    }
    ($chain | ForEach-Object { "{0:D10}" -f $_ }) -join '/'
}
$byId = @{}
foreach ($r in $rows) { $byId[[int]$r.Id] = $r }
$sorted = $rows | Sort-Object { Get-SortKey $_ $byId }

# --- 5. export ------------------------------------------------------------
if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
$sorted | Export-Excel -Path $OutFile -WorksheetName $SheetName `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableStyle Medium2

Write-Host "`nKlaar: $OutFile ($($sorted.Count) rijen)" -ForegroundColor Green
