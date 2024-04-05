param(
    [object]$WebhookData
)

# Function to validate an email address using regex
function Validate-EmailAddress {
    param([string]$email)
    $email -match "^\S+@\S+\.\S+$"
}

# Validate and cast input variables
if (-not $WebhookData.InputMailboxName -or $WebhookData.InputMailboxName -isnot [string]) {
    throw "InputMailboxName is not set or not a valid string."
}
if (-not $WebhookData.InputMailboxEmail -or $WebhookData.InputMailboxEmail -isnot [string]) {
    throw "InputMailboxEmail is not set or not a valid string."
}
if (-not $WebhookData.UsersToGrantAccess -or $WebhookData.UsersToGrantAccess -isnot [array]) {
    throw "UsersToGrantAccess is not set or not a valid string array."
}

$InputMailboxName = [string]$WebhookData.InputMailboxName
$InputMailboxEmail = [string]$WebhookData.InputMailboxEmail
$UsersToGrantAccess = [string[]]$WebhookData.UsersToGrantAccess

# Check if the email address is valid
if (-not (Validate-EmailAddress -email $InputMailboxEmail)) {
    throw "The email address $InputMailboxEmail is not a valid email format."
}

# Retrieve the stored credential from Azure Automation
$credential = Get-AutomationPSCredential -Name 'Srv_Automation_SharedMailboxCreation'

# Connect to Exchange Online using the credential
Connect-ExchangeOnline -Credential $credential -ShowProgress $true

# Check if the email address is already in use
if ($null -eq (Get-Recipient -Identity $InputMailboxEmail -ErrorAction SilentlyContinue)) {
    # Email address is not in use, so create the new shared mailbox
    New-Mailbox -Shared -Name $InputMailboxName -PrimarySmtpAddress $InputMailboxEmail
    
    # Assign full access to each user in the list
    foreach ($User in $UsersToGrantAccess) {
        if (-not (Validate-EmailAddress -email $User)) {
            throw "The user email address $User is not a valid email format."
        }
        Add-MailboxPermission -Identity $InputMailboxEmail -User $User -AccessRights FullAccess -InheritanceType All
        Add-RecipientPermission -Identity $InputMailboxEmail -Trustee $User -AccessRights SendAs -Confirm:$false
    }
} else {
    # Email address is in use, so output a message
    Write-Error "The email address $InputMailboxEmail is already in use."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false