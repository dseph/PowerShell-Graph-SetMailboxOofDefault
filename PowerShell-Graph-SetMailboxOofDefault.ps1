# PowerShell-Graph-SetMailboxOofDefault.ps1 
# This script demonstrates how to set default Out of Office (OOF) messages for a mailbox using Microsoft Graph API.
#
# It uses client credentials flow for authentication and updates the mailboxSettings for the specified user.    
# Usage: .\PowerShell-Graph-SetMailboxOofDefault.ps1 -Mailbox user@contoso
# Required permissions: MailboxSettings.ReadWrite (application permission with admin consent)   
# Note that this can be done in PowerShell also: 
#    https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxautoreplyconfiguration?view=exchange-ps
# Note this is an as-is sample for education. Test and make it your own before using in production and assume all responsiblities for its ussage.

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Mailbox
)

# Validate basic SMTP format
if ($Mailbox -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$") {
    Write-Error "Invalid SMTP address format: $Mailbox"
    exit
}

# =========================
# HARD-CODED CREDENTIALS
# =========================
$TenantId    = "YOUR_TENANT_ID"
$ClientId    = "YOUR_APP_ID"
$ClientSecret= "YOUR_APP_SECRET"

# =========================
# OOF SETTINGS (EDIT HERE)
# =========================
$InternalMessage = "<html><body><p>A default Internal auto-reply message.</p></body></html>"
$ExternalMessage = "<html><body><p>A default External auto-reply message.</p></body></html>"

# =========================
# GET ACCESS TOKEN
# =========================
$TokenRequestBody = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $ClientId
    client_secret = $ClientSecret
}

try {
    $TokenResponse = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
        -Body $TokenRequestBody

    $AccessToken = $TokenResponse.access_token
}
catch {
    Write-Error "Failed to acquire token: $_"
    exit
}

# =========================
# BUILD REQUEST BODY
# =========================
$Body = @{
    automaticRepliesSetting = @{
        status = "disabled"                     # alwaysEnabled | scheduled | disabled
        externalAudience = "all"                # none | contactsOnly | all
        internalReplyMessage = $InternalMessage
        externalReplyMessage = $ExternalMessage
    }
} | ConvertTo-Json -Depth 5

# =========================
# PATCH MAILBOX SETTINGS
# =========================
$Headers = @{
    Authorization = "Bearer $AccessToken"
    "Content-Type" = "application/json"
}

$Uri = "https://graph.microsoft.com/v1.0/users/$Mailbox/mailboxSettings"

try {
    Write-Host "Updating OOF for $Mailbox..."
    
    Invoke-RestMethod -Method PATCH `
        -Uri $Uri `
        -Headers $Headers `
        -Body $Body

    Write-Host "SUCCESS: OOF updated for $Mailbox"
}
catch {
    Write-Error "FAILED to update OOF: $_"
    if ($_.ErrorDetails.Message) {
        Write-Host $_.ErrorDetails.Message
    }
}
