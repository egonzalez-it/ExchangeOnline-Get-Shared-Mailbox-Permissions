<#
.SYNOPSIS
    Returns the shared mailboxes a user can access in Exchange Online.

.DESCRIPTION
    This script checks all shared mailboxes in Exchange Online and reports whether
    the specified user has Full Access, Send As, and/or Send on Behalf permissions.

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
$userRecipient = Get-Recipient -Identity $UserEmail

Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | ForEach-Object {
    $mailbox = $_
    $sharedMailbox = $mailbox.PrimarySmtpAddress

    $hasFullAccess = Get-MailboxPermission -Identity $sharedMailbox | Where-Object {
        $_.User -eq $UserEmail -and $_.AccessRights -contains "FullAccess"
    }

    $hasSendAs = Get-RecipientPermission -Identity $sharedMailbox | Where-Object {
        $_.Trustee -eq $UserEmail -and $_.AccessRights -contains "SendAs"
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
} | Sort-Object SharedMailbox | Format-Table -AutoSize