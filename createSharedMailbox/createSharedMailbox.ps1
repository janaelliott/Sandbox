param(
    [object]$WebhookData
)

$InputMailboxName = $WebhookData.InputMailboxName
$InputMailboxEmail = $WebhookData.InputMailboxEmail
$UsersToGrantAccess = $WebhookData.UsersToGrantAccess

# Retrieve the stored credential from Azure Automation
$credential = Get-AutomationPSCredential -Name 'Srv_Automation_SharedMailboxCreation'

# Connect to Exchange Online using the credential
Connect-ExchangeOnline -Credential $credential -ShowProgress $true

# Append (Skillsoft) to mailbox name
$SharedMailboxName = $InputMailboxName + " (Skillsoft)"

# If no email address is provided, create one based on the mailbox name
if (-not $InputMailboxEmail) {
    $StandardDomain = "@1b8v7r.onmicrosoft.com"
    # Remove characters that are not allowed in an email address
    $EmailLocalPart = $InputMailboxName -replace '[^\w.-]', '' # Keeps letters, numbers, periods, hyphens, and underscores
    $SharedMailboxEmail = $EmailLocalPart.ToLower() + $StandardDomain
} else {
    $SharedMailboxEmail = $InputMailboxEmail
}

# Check if the email address is already in use
if ($null -eq (Get-Recipient -Identity $SharedMailboxEmail -ErrorAction SilentlyContinue)) {
    # Email address is not in use, so create the new shared mailbox
    New-Mailbox -Shared -Name $SharedMailboxName -PrimarySmtpAddress $SharedMailboxEmail
    
    # Assign full access to each user in the list
    foreach ($User in $UsersToGrantAccess) {
        Add-MailboxPermission -Identity $SharedMailboxEmail -User $User -AccessRights FullAccess -InheritanceType All
        Add-RecipientPermission -Identity $SharedMailboxEmail -Trustee $User -AccessRights SendAs -Confirm:$false
    }
} else {
    # Email address is in use, so output a message
    Write-Host "The email address $SharedMailboxEmail is already in use."
}

# Disconnect from Exchange Online
Disconnect-ExchangeOnline -Confirm:$false