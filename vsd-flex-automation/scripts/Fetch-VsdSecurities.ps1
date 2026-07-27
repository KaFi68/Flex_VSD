<#
Crawls the full securities list ("Danh sách Chứng khoán") from vsd.vn (VSDC lookup tool)
and saves it as a timestamped JSON snapshot for later diffing.

Endpoint reverse-engineered from https://vsd.vn/vi/is (2026-07-27):
  - GET  https://vsd.vn/vi/is            -> session cookie + <meta name="__VPToken"> (CSRF token)
  - POST https://vsd.vn/isustocks/search  -> HTML table fragment, one page of results
      Body: { SearchKey, CurrentPage, RecordOnPage, OrderBy, OrderType }
      SearchKey format observed: "|<maCK>|<stockType>|<sanGiaoDich>|<status>|<lang>|||||"
      Header required: __VPToken must match the token issued in the same session
                        (double-submit pattern), else HTTP 400.
      NOTE: RecordOnPage is NOT honored by the server -- it always returns 10 rows/page,
            regardless of the value sent. Confirmed by testing 20/50/100/500.
#>

param(
    [string]$OutputDir = (Join-Path $PSScriptRoot "..\data\snapshots"),
    [int]$DelayMs = 300,          # polite delay between page requests
    [int]$MaxRetries = 3,
    [int]$MaxPages = 0            # 0 = crawl everything; >0 = stop early (for smoke testing)
)

$ErrorActionPreference = "Stop"

$BaseUrl   = "https://vsd.vn"
$ListPage  = "$BaseUrl/vi/is"
$SearchUrl = "$BaseUrl/isustocks/search"
$PageSize  = 10   # server-enforced, not actually configurable

function Get-VsdSession {
    $resp = Invoke-WebRequest -Uri $ListPage -SessionVariable session -UserAgent "Mozilla/5.0 (Windows NT 10.0; Win64; x64)" -UseBasicParsing
    $tokenMatch = [regex]::Match($resp.Content, '<meta name="__VPToken" content="([^"]*)"')
    if (-not $tokenMatch.Success) {
        throw "Khong tim thay __VPToken tren trang $ListPage - co the site da doi cau truc."
    }
    return [pscustomobject]@{
        Session = $session
        Token   = $tokenMatch.Groups[1].Value
    }
}

function Get-VsdSecuritiesPage {
    param($VsdSession, [int]$Page)

    $headers = @{
        "Content-Type"      = "application/json;charset=utf-8"
        "X-Requested-With"  = "XMLHttpRequest"
        "Referer"           = $ListPage
        "__VPToken"         = $VsdSession.Token
    }
    $body = @{
        SearchKey    = "|||||VI|||||"
        CurrentPage  = $Page
        RecordOnPage = $PageSize
        OrderBy      = "Code"
        OrderType    = "asc"
    } | ConvertTo-Json

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-WebRequest -Uri $SearchUrl -Method Post -WebSession $VsdSession.Session -Headers $headers -Body $body -UseBasicParsing
        } catch {
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Milliseconds (1000 * $attempt)
        }
    }
}

function Parse-VsdSecuritiesRows {
    param([string]$Html)

    $rows = [regex]::Matches($Html, '<tr>\s*(.*?)\s*</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($row in $rows) {
        $cells = [regex]::Matches($row.Groups[1].Value, '<td[^>]*>(.*?)</td>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($cells.Count -lt 8) { continue }  # skip header/malformed rows

        $cellText = $cells | ForEach-Object {
            $raw = $_.Groups[1].Value
            $stripped = [regex]::Replace($raw, '<[^>]+>', '').Trim()
            [System.Net.WebUtility]::HtmlDecode($stripped)
        }

        $results.Add([pscustomobject]@{
            Code              = $cellText[1]
            IsinCode          = $cellText[2]
            Name              = $cellText[3]
            StockType         = $cellText[4]
            Market            = $cellText[5]   # "Noi giao dich" - relevant for chuyen san
            ManagementArea    = $cellText[6]
            Status            = $cellText[7]
        })
    }
    return ,$results   # comma operator: giu nguyen la mot List, tranh bi PowerShell "unroll" thanh $null khi rong
}

function Get-TotalRecordCount {
    param([string]$Html)
    $m = [regex]::Match($Html, "1 - \d+ / (\d+) b" )
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

# ---- main ----

Write-Host "Dang khoi tao phien lam viec voi vsd.vn..."
$vsdSession = Get-VsdSession
Write-Host "Da lay duoc session + VPToken."

$firstPageResp = Get-VsdSecuritiesPage -VsdSession $vsdSession -Page 1
$total = Get-TotalRecordCount -Html $firstPageResp.Content
if (-not $total) { throw "Khong doc duoc tong so ban ghi tu trang dau tien." }
$totalPages = [math]::Ceiling($total / $PageSize)
if ($MaxPages -gt 0 -and $MaxPages -lt $totalPages) {
    Write-Host "Tong so ma chung khoan: $total ($totalPages trang) - gioi han test o $MaxPages trang"
    $totalPages = $MaxPages
} else {
    Write-Host "Tong so ma chung khoan: $total ($totalPages trang)"
}

$all = New-Object System.Collections.Generic.List[object]
$all.AddRange((Parse-VsdSecuritiesRows -Html $firstPageResp.Content))

for ($p = 2; $p -le $totalPages; $p++) {
    Start-Sleep -Milliseconds $DelayMs
    $resp = Get-VsdSecuritiesPage -VsdSession $vsdSession -Page $p
    $rows = Parse-VsdSecuritiesRows -Html $resp.Content

    if ($rows.Count -eq 0) {
        Write-Warning "Trang $p tra ve 0 dong, thu lai 1 lan..."
        Start-Sleep -Milliseconds 1000
        $resp = Get-VsdSecuritiesPage -VsdSession $vsdSession -Page $p
        $rows = Parse-VsdSecuritiesRows -Html $resp.Content
        if ($rows.Count -eq 0) {
            Write-Warning "Trang $p van tra ve 0 dong sau khi thu lai - bo qua trang nay."
        }
    }

    $all.AddRange($rows)

    if ($p % 20 -eq 0 -or $p -eq $totalPages) {
        Write-Host "  ... da quet trang $p / $totalPages ($($all.Count) ma)"
    }
}

if ($all.Count -ne $total) {
    Write-Warning "So ban ghi thu thap duoc ($($all.Count)) khac voi tong bao cao ($total). Co the co trang bi loi hoac du lieu thay doi giua luc quet."
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutputDir "vsd_securities_$timestamp.json"
$all | ConvertTo-Json -Depth 3 | Out-File -FilePath $outFile -Encoding utf8

Write-Host "Da luu snapshot: $outFile ($($all.Count) ma chung khoan)"
