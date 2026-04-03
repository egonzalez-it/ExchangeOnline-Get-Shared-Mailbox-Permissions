<#
.SYNOPSIS
    Returns the shared mailboxes a user can access in Exchange Online.

.DESCRIPTION
    This script checks all shared mailboxes in Exchange Online and reports whether
    the specified user has Full Access, Send As, and/or Send on Behalf permissions.
    It displays the results on screen once and exports them to:
    SharedMailboxAccess_<User>_<YYYY-MM-DD>.csv
    in the same folder where the script is located.

.PARAMETER UserEmail
    The email address of the user to check.

.EXAMPLE
    .\Get-SharedMailboxAccess.ps1 -UserEmail user@company.com
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserEmail
)

# Resolve the user first
$userRecipient = Get-Recipient -Identity $UserEmail -ErrorAction Stop

# Build output folder in the same folder as the script
$scriptFolder = if ($PSScriptRoot) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

# Clean username for file name
$userNameClean = ($UserEmail -split "@")[0]

# Get current date
$dateStamp = Get-Date -Format "yyyy-MM-dd"

# Build output file name
$outputFile = Join-Path -Path $scriptFolder -ChildPath "SharedMailboxAccess_${userNameClean}_${dateStamp}.csv"

# Collect results
$results = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_
    $sharedMailbox = $mailbox.PrimarySmtpAddress.ToString()

    $hasFullAccess = Get-MailboxPermission -Identity $sharedMailbox -ErrorAction SilentlyContinue | Where-Object {
        $_.User -eq $UserEmail -and $_.AccessRights -contains 'FullAccess'
    }

    $hasSendAs = Get-RecipientPermission -Identity $sharedMailbox -ErrorAction SilentlyContinue | Where-Object {
        $_.Trustee -eq $UserEmail -and $_.AccessRights -contains 'SendAs'
    }

    $hasSendOnBehalf = $mailbox.GrantSendOnBehalfTo | Where-Object {
        $_ -eq $userRecipient.Name -or
        $_ -eq $userRecipient.Alias -or
        $_ -eq $userRecipient.DistinguishedName
    }

    if ($hasFullAccess -or $hasSendAs -or $hasSendOnBehalf) {
        [PSCustomObject]@{
            SharedMailbox = $sharedMailbox
            FullAccess    = [bool]$hasFullAccess
            SendAs        = [bool]$hasSendAs
            SendOnBehalf  = [bool]$hasSendOnBehalf
        }
    }
} | Sort-Object SharedMailbox

# Show results on screen once and export
if ($results) {
    $results | Format-Table -AutoSize
    $results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Results exported to: $outputFile" -ForegroundColor Green
}
else {
    Write-Host "No shared mailbox permissions found for $UserEmail." -ForegroundColor Yellow
}
