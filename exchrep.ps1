<#
Universal Exchange Server report. This PowerShell script connects to Exchange Online (or optionally on-prem Exchange),
runs predefined PowerShell queries (mailboxes, contacts, mailbox database copy status, etc.), generates HTML reports with optional charts,
compresses them into ZIP files, sends the report via SMTP or Microsoft Graph, optionally posts a message to Microsoft Teams,
and can display a dashboard that refreshes automatically.

Andrzej Sulich 2025
#>

param(
    [switch]$Dashboard,
    [int]$Refresh = 2
)

# ========== CONFIG ==========
$OutputDir     = "$PSScriptRoot\Reports"
$LogFile       = "$PSScriptRoot\report.log"	
$EnableCharts  = $True				# Enable Charts 
$EnableZip     = $false   			# Enable reports zip archiving
$EnableTeams   = $false   			# automatic reports on Teams
$TeamsWebhook  = "https://YOUR_TEAMS_WEBHOOK"

# SMTP (jeśli używasz SMTP)
$EmailFrom     = "report@domain.com"
$EmailTo       = "admin@domain.com"
$EmailSubject  = "Exchange Report"
$SmtpServer    = "smtp.office365.com"
$SmtpPort      = 587
$SmtpUser      = "report@domain.com"
$SmtpPass      = "APP_PASSWORD_HERE"   # jeśli konto ma MFA → użyj Graph API

# Graph API (alternatywa dla MFA)
$UseGraph      = $false
$GraphTenantId = ""
$GraphClientId = ""
$GraphSecret   = ""

# ============================

# ========== LOGGING ==========
function Write-Log {
    param([string]$Message,[string]$Level="INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File $LogFile -Append
}

Write-Log "Script started"

# Create output directory
if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# ========== CONNECT EXCHANGE ONLINE ==========
function Connect-ExchangeCloud {
    try {
        Write-Log "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Connected"
    } catch {
        Write-Log "Failed to connect EXO: $_" "ERROR"
    }
}

# ========== RUN QUERY ==========
function Run-Query {
    param(
        [string]$Name,
        [string]$ScriptBlock
    )

    Write-Log "Running query: $Name"

    try {
        $data = Invoke-Expression $ScriptBlock
        return [PSCustomObject]@{
            Name = $Name
            Data = $data
        }
    } catch {
        Write-Log "Query failed: $Name - $_" "ERROR"
        return [PSCustomObject]@{
            Name = $Name
            Data = @()
        }
    }
}

# ========== MAIL SENDERS ==========
function Send-MailSMTP {
    param([string]$AttachmentPath)

    try {
        Send-MailMessage `
            -From $EmailFrom `
            -To $EmailTo `
            -Subject $EmailSubject `
            -Body "Report generated. HTML attached." `
            -Attachments $AttachmentPath `
            -SmtpServer $SmtpServer `
            -Port $SmtpPort `
            -UseSsl `
            -Credential (New-Object System.Management.Automation.PSCredential($SmtpUser,(ConvertTo-SecureString $SmtpPass -AsPlainText -Force)))

        Write-Log "Email sent (SMTP)"
    } catch {
        Write-Log "SMTP send failed: $_" "ERROR"
    }
}

function Send-MailGraph {
    Write-Log "Sending email via Graph API (configure yourself)"
}

# ========== TEAMS ==========
function Send-TeamsMessage {
    param([string]$Message)

    if (-not $EnableTeams) { return }

    try {
        Invoke-RestMethod -Method POST -Uri $TeamsWebhook `
            -ContentType "application/json" `
            -Body (@{ text = $Message } | ConvertTo-Json)
        Write-Log "Teams message delivered"
    } catch {
        Write-Log "Teams failed: $_" "ERROR"
    }
}

# ========== GENERATE HTML ==========
function Generate-HTML {
    param(
        [array]$Results
    )

    $date = Get-Date -Format "yyyyMMdd_HHmm"
    $file = Join-Path $OutputDir "ExchangeReport_$date.html"

    $html = @"
<html>
<head>
<meta charset="UTF-8">
<title>Exchange Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body { font-family: Arial; background:#111; color:#eee; padding:20px; }
table { border-collapse:collapse; width:100%; margin:20px 0; }
th,td { border:1px solid #555; padding:6px; }
th { background:#333; }
.chartwrap { margin:40px 0; background:#222; padding:20px; border-radius:10px; }
</style>
</head>
<body>
<h1>Exchange Report</h1>
<p>Generated: $(Get-Date)</p>
"@

    foreach ($item in $Results) {
        $html += "<h2>$($item.Name)</h2>"

        if ($item.Data.Count -eq 0) {
            $html += "<i>No data</i>"
            continue
        }

        # Table
        $html += ($item.Data | ConvertTo-Html -Fragment)

        # Numeric properties
        $numericProps = ($item.Data | Select-Object -First 20 | ForEach-Object {
            $_ | Get-Member -MemberType NoteProperty |
                Where-Object {
                    $_.Definition -match "Int32" -or
                    $_.Definition -match "Int64" -or
                    $_.Definition -match "Double" -or
                    $_.Definition -match "Decimal"
                } | Select-Object -ExpandProperty Name
        }) | Select-Object -Unique

        if ($EnableCharts -and $numericProps.Count -gt 0) {
            $first = $numericProps[0]

            $labels = $item.Data | ForEach-Object {
                if ($_.PSObject.Properties["DisplayName"]) { $_.DisplayName }
                elseif ($_.PSObject.Properties["Name"]) { $_.Name }
                else { "Item" }
            }

            $values = $item.Data | ForEach-Object { $_.$first }

            $chartId = "chart_" + ([guid]::NewGuid().ToString("N"))

            $html += "<div class='chartwrap'><canvas id='$chartId'></canvas></div>"
            $html += "<script>
var ctx = document.getElementById('$chartId').getContext('2d');
new Chart(ctx,{
 type:'bar',
 data:{ labels: $(ConvertTo-Json $labels), datasets:[{ label:'$first', data: $(ConvertTo-Json $values) }] },
 options:{ responsive:true }
});
</script>"
        }
    }

    $html += "</body></html>"
    $html | Out-File $file -Encoding UTF8

    return $file
}

# ========== ZIP ==========
function Zip-Reports {
    if (-not $EnableZip) { return $null }

    $zipFile = Join-Path $OutputDir ("Reports_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".zip")

    if (Test-Path $zipFile) { Remove-Item $zipFile -Force }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        # Create empty ZIP
        $zip = [System.IO.Compression.ZipFile]::Open($zipFile, [System.IO.Compression.ZipArchiveMode]::Create)

        # Add only HTML files
        Get-ChildItem $OutputDir -Filter *.html | ForEach-Object {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $_.FullName, $_.Name)
        }

        $zip.Dispose()
        Write-Log "ZIP created with HTML files only: $zipFile"
        return $zipFile
    } catch {
        Write-Log "ZIP failed: $_" "ERROR"
        return $null
    }
}


# ========== DASHBOARD ==========
function Start-Dashboard {
    param([int]$RefreshMinutes = 2)

    Write-Log "Dashboard started with refresh: $RefreshMinutes min"

    while ($true) {

        $dash = foreach ($q in $Queries) {
            Run-Query -Name $q.Name -ScriptBlock $q.Query
        }

        $html = @"
<html>
<head>
<meta charset="UTF-8">
<title>Exchange Dashboard</title>
<style>
body { background:#222; color:#eee; font-family:Arial; }
</style>
</head>
<body>
<h1>Exchange Dashboard</h1>
<p>Updated: $(Get-Date)</p>
"@

        foreach ($item in $dash) {
            $html += "<h2>$($item.Name)</h2>"
            if ($item.Data.Count -eq 0) {
                $html += "<i>No data</i>"
            } else {
                $html += ($item.Data | ConvertTo-Html -Fragment)
            }
        }

        $html += "</body></html>"

        $file = "$OutputDir\Dashboard.html"
        $html | Out-File $file -Encoding UTF8
        Start-Process $file

        Start-Sleep -Seconds ($RefreshMinutes * 60)
    }
}

# ========== DEFINE QUERY SET ==========
$Queries = @(
    @{ Name="MailboxList"     ; Query="Get-Mailbox -ResultSize 50" }
    @{ Name="Contacts"        ; Query="Get-Contact -ResultSize 50" }
    @{ Name="DBCopyStatus"    ; Query="Get-MailboxDatabaseCopyStatus *" }
    @{ Name="Mailbox"         ; Query="Get-MailboxDatabase DB01 | select Name, ServerName, OriginatingServer, MetaCacheDatabaseRootFolderPath, MetaCacheDatabaseMountpointFolderPath, MetaCacheDatabaseFolderPath, MetaCacheDatabaseFilePath, MetaCacheDatabaseMaxCapacityInBytes" }
)

# ========== EXECUTION FLOW ==========

Connect-ExchangeCloud

if ($Dashboard) {
    Start-Dashboard -RefreshMinutes $Refresh
    exit
}

$results = foreach ($q in $Queries) {
    Run-Query -Name $q.Name -ScriptBlock $q.Query
}

$htmlFile = Generate-HTML -Results $results

$zipFile = Zip-Reports

Send-TeamsMessage "Exchange report completed"

if ($UseGraph) { Send-MailGraph }
else { Send-MailSMTP -AttachmentPath $zipFile }

Write-Log "Script completed"
