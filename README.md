**# Exchange-Server-Reporting-Script**

**Script Name:**
    exchrep.ps1

**Description:**
    This PowerShell script connects to Exchange Online (or optionally on-prem Exchange), runs predefined PowerShell queries (mailboxes, contacts, mailbox database copy status, etc.), generates HTML reports with optional charts, compresses them into ZIP files, sends the report via SMTP or Microsoft Graph, optionally posts a message to Microsoft Teams, and can display a dashboard that refreshes automatically.

**Parameters:**
    Parameter	    Type	    Default	  Description
      -Dashboard	[switch]	$false	  If specified, runs the script in dashboard mode, which continuously refreshes and displays query results in an HTML page.
      -Refresh	  [int]	    2	        Used with -Dashboard to define refresh interval in minutes for the dashboard.

**Variables:**
    Parameter	      Type	      Default	                        Description
    $EnableCharts	  [bool]	    $true	                          Enables chart generation in HTML reports for numeric properties.
    $EnableZip	    [bool]	    $true	                          Enables creation of ZIP file containing the report(s).
    $EnableTeams	  [bool]	    $false	                        Enables sending a message to a Teams channel via webhook.
    $TeamsWebhook	  [string]	  "https://YOUR_TEAMS_WEBHOOK"	  URL of Teams webhook if $EnableTeams is $true.
    $OutputDir	    [string]	  "$PSScriptRoot\Reports"	        Directory to save HTML reports and ZIP files.
    $LogFile	      [string]	  "$PSScriptRoot\report.log"	    File where script logs actions and errors.
    $EmailFrom	    [string]	  "report@domain.com"	            Sender email address for SMTP/Graph email.
    $EmailTo	      [string]	  "admin@domain.com"	            Recipient email address.
    $EmailSubject	  [string]	  "Exchange Report"	              Email subject line.
    $SmtpServer	    [string]	  "smtp.office365.com"	          SMTP server for sending emails.
    $SmtpPort	      [int]	      587	                            SMTP port (typically 587 for TLS).
    $SmtpUser	      [string]	  "report@domain.com"	            SMTP username (email account).
    $SmtpPass	      [string]	  "APP_PASSWORD_HERE"	            SMTP password or app password.
    $UseGraph	      [bool]	    $false	                        Use Microsoft Graph API instead of SMTP for sending emails (required for MFA accounts).
    $GraphTenantId	[string]	  ""	                            Tenant ID for Graph API authentication.
    $GraphClientId	[string]	  ""	                            App client ID for Graph API.
    $GraphSecret	  [string]	  ""	                            App secret for Graph API.

**Functions Overview**
Function								Description
Connect-ExchangeCloud		Connects to Exchange Online.
Run-Query								Executes a PowerShell query and returns a custom object containing name and results.
Generate-HTML						Generates an HTML report from query results, with optional charts for numeric properties.
Zip-Reports							Creates a ZIP file containing only HTML reports in $OutputDir.
Send-MailSMTP						Sends the report via SMTP with attachments.
Send-MailGraph					Placeholder for sending email using Graph API (for MFA-enabled accounts).
Send-TeamsMessage				Posts a message to a Teams channel via webhook.
Start-Dashboard					Runs the dashboard mode, refreshes every -Refresh minutes, and displays HTML query results continuously.
Write-Log								Logs actions and errors to $LogFile.

**Predefined Queries**
The script defines these queries by default:
Name					PowerShell Command								Notes
MailboxList		Get-Mailbox -ResultSize 50				Retrieves 50 mailboxes.
Contacts			Get-Contact -ResultSize 50				Retrieves 50 contacts.
DBCopyStatus	Get-MailboxDatabaseCopyStatus *		Shows copy status for all mailbox databases.

Additional queries can be added to the $Queries array:
_$Queries += @{ Name="TransportQueues"; Query="Get-Queue" }_

------------------------------------------------------------

**Usage Examples**

**1. Run standard report and send via SMTP**
_.\exchrep.ps1_
Generates HTML report(s) in Reports folder.
Creates ZIP file containing only HTML files.
Sends report via SMTP.

**2. Run report and send to Teams**
_$EnableTeams = $true
$TeamsWebhook = "https://outlook.office.com/webhook/..."
.\exchrep.ps1_
Sends a Teams notification after the report is generated.

**3. Run dashboard mode**
_.\exchrep.ps1 -Dashboard_
Opens an HTML dashboard that auto-refreshes every 2 minutes (default).
_.\exchrep.ps1 -Dashboard -Refresh 5_
Opens a dashboard that refreshes every 5 minutes.

**4. Enable/Disable charts**
_$EnableCharts = $false
.\exchrep.ps1_
Charts are disabled.

**5. Enable/Disable ZIP**
_$EnableZip = $false
.\exchrep.ps1_
Skips ZIP creation, only HTML files remain.

**6. Use Graph API for MFA accounts**
_$UseGraph      = $true
$GraphTenantId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$GraphClientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$GraphSecret   = "SECRET_VALUE"
.\exchrep.ps1_
Uses Graph API to send email, bypassing SMTP issues with MFA-enabled accounts.

**7. Adding a custom query**
_$Queries += @{ Name="TransportQueues"; Query="Get-Queue" }_
The script will generate a separate table (and chart if numeric) for this query.

------------------------------------------------------------

**Notes / Recommendations**

Exchange Connection
	- Ensure Exchange Online module is installed:
	- Install-Module -Name ExchangeOnlineManagement
	- Connect manually if needed: Connect-ExchangeOnline.

SMTP and MFA
	- Standard Send-MailMessage fails with MFA.
	- Use app password or switch to Graph API for MFA accounts.

Dashboard Mode
	- Runs in a loop until the script is stopped.
	- Refresh interval can be adjusted with -Refresh.

HTML & Charts
	- Charts only generated if numeric data is detected in query results.

ZIP File
	- Only HTML files from $OutputDir are included.

Logs
	- All script actions and errors are logged to $LogFile.
