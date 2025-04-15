# ========================== #
#     Exchange Online Login  #
#     via App Registration   #
# ========================== #

# =============================== #
#     Initial Setup für Skript    #
#     Verzeichnis + Config Check  #
# =============================== #

# Define folder paths
$logDir = "$PSScriptRoot\logs"
$configDir = "$PSScriptRoot\config"
$configFile = "$configDir\config.json"

# Create logs directory if it doesn't exist
if (-not (Test-Path -Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    Write-Host "Created logs directory at: $logDir"
}

# Create config directory if it doesn't exist
if (-not (Test-Path -Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    Write-Host "Created config directory at: $configDir"
}

# Create default config.json if it doesn't exist
if (-not (Test-Path -Path $configFile)) {
    $defaultConfig = @{
        AppId = "your-app-id"
        CertificateThumbprint = "your-certificate-thumbprint"
        Organization = "your-organization.onmicrosoft.com"
        TenantId = "your-tenant-id"
        ClientSecret = "your-client-secret"
        FromEmail = "your-sender@example.com"
        ToEmail = "your-recipient@example.com"
        CCEmail = "your-cc@example.com"
    } | ConvertTo-Json -Depth 3

    $defaultConfig | Set-Content -Path $configFile -Encoding UTF8
    Write-Host "Created example config.json at: $configFile"
    Write-Host "`nPlease update the config.json with your actual credentials and email addresses.`n"
}


# Define log file path
$logFile = "$PSScriptRoot\logs\ERA_Report_Log.txt"

# Function for writing logs
function Write-Log {
    param (
        [string]$message,
        [string]$logFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logMessage
}

# Start of logging
Write-Log -message "Script started." -logFile $logFile

# Read JSON file and convert it into a PowerShell object
try {
    $configPath = "$PSScriptRoot\config\config.json"
    $config = Get-Content -Path $configPath | ConvertFrom-Json
    Write-Log -message "Configuration file successfully loaded." -logFile $logFile
} catch {
    Write-Log -message "Error loading configuration file: $_" -logFile $logFile
    exit 1
}

# Extract configuration values from JSON file
$AppId = $config.AppId
$CertificateThumbprint = $config.CertificateThumbprint
$Organization = $config.Organization
$tenantId = $config.TenantId
$clientSecret = $config.ClientSecret
$from = $config.FromEmail
$to = $config.ToEmail
$cc = $config.CCEmail

# Format report date
$reportDate = Get-Date -Format "MM-yyyy"

# Connect to Exchange Online via App (Client Credentials Flow)
try {
    Write-Log -message "Attempting to connect to Exchange Online." -logFile $logFile
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    Write-Log -message "Successfully connected to Exchange Online." -logFile $logFile
} catch {
    Write-Log -message "Error connecting to Exchange Online: $_" -logFile $logFile
    exit 1
}

# =============================== #
#     Count all mailboxes         #
# =============================== #

try {
    Write-Log -message "Counting all mailboxes." -logFile $logFile
    # Count user mailboxes and shared mailboxes
    $userMailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
    $sharedMailboxes = Get-ExoMailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

    # Total number of mailboxes
    $totalMailboxes = ($userMailboxes + $sharedMailboxes).Count
    Write-Log -message "Mailbox count: User Mailboxes = $($userMailboxes.Count), Shared Mailboxes = $($sharedMailboxes.Count), Total = $totalMailboxes." -logFile $logFile
} catch {
    Write-Log -message "Error counting mailboxes: $_" -logFile $logFile
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# =============================== #
#     Send email (Graph API)      #
# =============================== #

$subject = "ERA-Report ORGANIZATION $reportDate"

$bodyText = @"
Hello,<br><br>

Exchange Online mailbox statistics (ERA billing):<br><br>

Number of <b>user mailboxes</b>: $($userMailboxes.Count)<br>
Number of <b>shared mailboxes</b>: $($sharedMailboxes.Count)<br>
<b>Total number</b> of mailboxes: <b>$totalMailboxes</b>
<br><br>
Kind regards<br><br>
Exchange Automation Bot
"@

try {
    # 1. Get access token
    Write-Log -message "Retrieving access token for Graph API." -logFile $logFile
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $tokenBody = @{
        client_id     = $AppId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody
    $accessToken = $tokenResponse.access_token
    Write-Log -message "Access token successfully retrieved." -logFile $logFile

    # 2. Create JSON for the email
    $mailBody = @{
        message = @{
            subject = $subject
            body = @{
                contentType = "HTML"
                content = $bodyText
            }
            toRecipients = @(@{
                emailAddress = @{
                    address = $to
                }
            })
            ccRecipients = @(@{
                emailAddress = @{
                    address = $cc
                }
            })
        }
        saveToSentItems = "false"
    } | ConvertTo-Json -Depth 10

    # 3. Use Graph API: send as $from
    $uri = "https://graph.microsoft.com/v1.0/users/$from/sendMail"

    Write-Log -message "Sending email via Graph API." -logFile $logFile
    Invoke-RestMethod -Method Post -Uri $uri -Headers @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type"  = "application/json"
    } -Body $mailBody -ContentType "application/json; charset=utf-8"

    Write-Log -message "Email successfully sent." -logFile $logFile
} catch {
    Write-Log -message "Error sending email via Graph API: $_" -logFile $logFile
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Disconnect from Exchange Online
try {
    Write-Log -message "Disconnecting from Exchange Online." -logFile $logFile
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Log -message "Successfully disconnected." -logFile $logFile
} catch {
    Write-Log -message "Error disconnecting from Exchange Online: $_" -logFile $logFile
}

Write-Log -message "Script finished." -logFile $logFile