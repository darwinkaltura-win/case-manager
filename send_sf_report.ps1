# Salesforce Team Open Cases Report
# Sends HTML email via Outlook

param(
    [string]$To = "darwin.mitra@kaltura.com",
    [string]$Subject = "Team Open Cases Report - $(Get-Date -Format 'MM/dd/yyyy')"
)

# ── CONFIG ────────────────────────────────────────────────────────────────────
$sfOrg = "kaltura"
$teamNames = @(
    'Russ Lichterman', 'Darwin Mitra', 'Fahad Mizi', 'Alex De Los Santos',
    'Tahmid Hassan', 'Roxy Hennessy', 'Rick Rehmann', 'Zach Hill',
    'Oscar Lagua Espin', 'Hector Zurita', 'Agustin Herling', 'Stivan Tenev',
    'Asad Ali', 'Julian Lucena Herrera', 'Renato Pinheiro'
)
$displayNames = @{
    'Oscar Lagua Espin'   = 'Oscar Espin'
    'Julian Lucena Herrera' = 'Julian Herrera'
}
$openStatuses = @(
    'New','In Progress','In Work','Awaiting Customer Response','Awaiting CSM',
    'Awaiting Tier 3','Awaiting ADA Team','Awaiting PS','Awaiting Product',
    'Awaiting Vendor','Awaiting R&D','Awaiting Internal','Awaiting FR Review',
    'Awaiting Deployment','Awaiting DevOps','Awaiting Owner Response',
    'Awaiting Response','Awaiting Internal Email','Review Customer Response',
    'Review Internal','Review JIRA Response','Review Akamai Response',
    'Review Customer Response (Reopened)','On Hold','FR In Review',
    'Customer Responded','Will be closed in 48H','Submitted','Resource Requested',
    'Resolved','Solution Provided to Customer',
    'Recommend to Close - Solution Provided',
    'Recommend to Close - No Longer Needed','Approved by Manager'
)
$actionableStatuses = @(
    'New','In Progress','In Work','Customer Responded',
    'Review Customer Response','Review Customer Response (Reopened)',
    'Review Internal','Review JIRA Response','Review Akamai Response',
    'Resolved','FR In Review','Will be closed in 48H',
    'Recommend to Close - Solution Provided',
    'Recommend to Close - No Longer Needed','Approved by Manager'
)
$statusInClause = ($openStatuses | ForEach-Object { "'$_'" }) -join ','
$nameInClause   = ($teamNames   | ForEach-Object { "'$_'" }) -join ','

# ── QUERY SALESFORCE ──────────────────────────────────────────────────────────
Write-Host "Querying Salesforce..." -ForegroundColor Cyan

$result = sf data query --target-org $sfOrg --result-format csv --query `
    "SELECT CaseNumber, Subject, Status, Priority, Owner.Name, IsEscalated, FLAGS__Case_Flags_Sort__c FROM Case WHERE Owner.Name IN ($nameInClause) AND Status IN ($statusInClause) ORDER BY Owner.Name, CreatedDate DESC LIMIT 2000" 2>&1

$lines = $result | Where-Object { $_ -match '^[0-9]{8},' -or $_ -match '^CaseNumber' }
$tmpCsv = "$env:TEMP\sf_report_temp.csv"
$lines | Out-File -Encoding utf8 $tmpCsv
$data = Import-Csv $tmpCsv

Write-Host "  $($data.Count) open cases retrieved." -ForegroundColor Green

# ── COMPUTE COUNTS ────────────────────────────────────────────────────────────
$summary = foreach ($name in $teamNames) {
    $cases      = @($data | Where-Object { $_.'Owner.Name' -eq $name -and $_.CaseNumber -match '^\d' })
    $open       = $cases.Count
    $escalated  = @($cases | Where-Object { $_.IsEscalated -eq 'true' }).Count
    $actionable = @($cases | Where-Object { $actionableStatuses -contains $_.Status }).Count
    $escAction  = @($cases | Where-Object { $_.IsEscalated -eq 'true' -and $actionableStatuses -contains $_.Status }).Count
    $blackFlag  = @($cases | Where-Object { $_.'FLAGS__Case_Flags_Sort__c' -like 'L4*' }).Count
    $display    = if ($displayNames[$name]) { $displayNames[$name] } else { $name }
    [PSCustomObject]@{
        Name              = $display
        SfName            = $name
        'Open Tickets'    = $open
        Escalated         = $escalated
        Actionables       = $actionable
        'Esc. Actionables'= $escAction
        'Black Flags'     = $blackFlag
    }
}

$escActionCases = $data | Where-Object {
    $_.CaseNumber -match '^\d' -and
    $_.IsEscalated -eq 'true' -and
    $actionableStatuses -contains $_.Status
} | Sort-Object 'Owner.Name'

$blackFlagCases = $data | Where-Object {
    $_.CaseNumber -match '^\d' -and
    $_.'FLAGS__Case_Flags_Sort__c' -like 'L4*'
} | Sort-Object 'Owner.Name'

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
function th($text) { "<th>$text</th>" }
function td($text, $align = 'left') { "<td style='text-align:$align'>$text</td>" }
function tdn($n) { if ($n -eq 0) { "<td style='text-align:center;color:#aaa'>0</td>" } else { "<td style='text-align:center;font-weight:bold'>$n</td>" } }

function Get-DisplayName($sfName) {
    if ($displayNames.ContainsKey($sfName)) { return $displayNames[$sfName] }
    return $sfName
}

# ── BUILD HTML ────────────────────────────────────────────────────────────────
$runTime = Get-Date -Format "dddd, MMMM d yyyy 'at' h:mm tt"

$html = @"
<html><head><style>
  body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #222; }
  h2   { color: #0078d4; border-bottom: 2px solid #0078d4; padding-bottom: 4px; }
  h3   { color: #444; margin-top: 28px; }
  table { border-collapse: collapse; margin-bottom: 24px; min-width: 500px; }
  th   { background: #0078d4; color: #fff; padding: 7px 12px; text-align: left; font-size: 12px; }
  td   { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; font-size: 12px; }
  tr:hover td { background: #f0f6ff; }
  .total td { background: #f5f5f5; font-weight: bold; border-top: 2px solid #0078d4; }
  .flag { color: #b00; font-weight: bold; }
  .esc  { color: #e65c00; font-weight: bold; }
  .zero { color: #aaa; }
  .ts  { font-size: 11px; color: #888; margin-bottom: 20px; }
</style></head><body>

<h2>Team Open Cases Report</h2>
<p class='ts'>Generated: $runTime</p>

<h3>Table 1 — Summary</h3>
<table>
<tr>$(th 'Name')$(th 'Open Tickets')$(th 'Escalated')$(th 'Actionables')$(th 'Esc. Actionables')$(th 'Black Flags')</tr>
"@

foreach ($row in $summary) {
    $html += "<tr>"
    $html += td $row.Name
    $html += tdn $row.'Open Tickets'
    $html += tdn $row.Escalated
    $html += tdn $row.Actionables
    $html += tdn $row.'Esc. Actionables'
    $html += tdn $row.'Black Flags'
    $html += "</tr>`n"
}

$totOpen  = ($summary | Measure-Object 'Open Tickets' -Sum).Sum
$totEsc   = ($summary | Measure-Object Escalated -Sum).Sum
$totAct   = ($summary | Measure-Object Actionables -Sum).Sum
$totEscA  = ($summary | Measure-Object 'Esc. Actionables' -Sum).Sum
$totBF    = ($summary | Measure-Object 'Black Flags' -Sum).Sum

$html += "<tr class='total'><td>TOTAL</td><td style='text-align:center'>$totOpen</td><td style='text-align:center'>$totEsc</td><td style='text-align:center'>$totAct</td><td style='text-align:center'>$totEscA</td><td style='text-align:center'>$totBF</td></tr>"
$html += "</table>"

# Table 2 — Escalated by Owner
$html += "<h3>Table 2 - Escalated Cases by Owner</h3><table>"
$html += "<tr>$(th 'Owner')$(th 'Escalated Cases')</tr>"
foreach ($row in ($summary | Where-Object { $_.Escalated -gt 0 } | Sort-Object Escalated -Descending)) {
    $html += "<tr>$(td $row.Name)<td style='text-align:center;font-weight:bold' class='esc'>$($row.Escalated)</td></tr>`n"
}
$html += "<tr class='total'><td>TOTAL</td><td style='text-align:center'>$totEsc</td></tr></table>"

# Table 3 — Black Flags by Owner
$html += "<h3>Table 3 - Black Flag Cases by Owner</h3><table>"
$html += "<tr>$(th 'Owner')$(th 'Black Flag Cases')</tr>"
foreach ($row in ($summary | Where-Object { $_.'Black Flags' -gt 0 } | Sort-Object 'Black Flags' -Descending)) {
    $html += "<tr>$(td $row.Name)<td style='text-align:center;font-weight:bold' class='flag'>$($row.'Black Flags')</td></tr>`n"
}
$html += "<tr class='total'><td>TOTAL</td><td style='text-align:center'>$totBF</td></tr></table>"

# Table 4 — Escalated Actionables detail
$escActCount = @($escActionCases).Count
$html += "<h3>Table 4 - Escalated Actionables ($escActCount cases)</h3>"
if (@($escActionCases).Count -eq 0) {
    $html += "<p style='color:#888'>No escalated actionable cases.</p>"
} else {
    $html += "<table><tr>$(th 'Case #')$(th 'Owner')$(th 'Subject')$(th 'Status')$(th 'Priority')</tr>"
    foreach ($c in $escActionCases) {
        $dname = Get-DisplayName $c.'Owner.Name'
        $html += "<tr>$(td $c.CaseNumber)$(td $dname)$(td $c.Subject)$(td $c.Status)$(td $c.Priority)</tr>`n"
    }
    $html += "</table>"
}

# Table 5 — Black Flag detail
$bfCount = @($blackFlagCases).Count
$html += "<h3>Table 5 - Black Flag Cases ($bfCount cases)</h3>"
if (@($blackFlagCases).Count -eq 0) {
    $html += "<p style='color:#888'>No black flag cases.</p>"
} else {
    $html += "<table><tr>$(th 'Case #')$(th 'Owner')$(th 'Subject')$(th 'Status')$(th 'Priority')</tr>"
    foreach ($c in $blackFlagCases) {
        $dname = Get-DisplayName $c.'Owner.Name'
        $html += "<tr>$(td $c.CaseNumber)$(td $dname)$(td $c.Subject)$(td $c.Status)$(td $c.Priority)</tr>`n"
    }
    $html += "</table>"
}

$html += "</body></html>"

# ── SEND VIA OUTLOOK ──────────────────────────────────────────────────────────
Write-Host "Sending email to $To..." -ForegroundColor Cyan

$outlook = New-Object -ComObject Outlook.Application
$mail    = $outlook.CreateItem(0)

$mail.To      = $To
$mail.Subject = $Subject
$mail.HTMLBody = $html

$mail.Send()

Write-Host "Email sent successfully to $To" -ForegroundColor Green
