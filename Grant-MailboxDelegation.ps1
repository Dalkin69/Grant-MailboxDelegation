<#
.SYNOPSIS
    Grants Full Access and Send As permissions on a list of Exchange Online mailboxes to a single delegate.

.DESCRIPTION
    Reads a text file of mailbox UPNs (one per line) and grants the delegate FullAccess and SendAs
    on each mailbox. Existing permissions are skipped. See README.md for full documentation.

.PARAMETER Delegate
    UPN of the user who will RECEIVE the permissions.

.PARAMETER MailboxListPath
    Path to a text file containing the UPNs of the mailboxes to delegate (one per line).

.PARAMETER AutoMapping
    If $true (default), mailboxes are automatically added to the delegate's Outlook profile.

.PARAMETER ReportPath
    Optional path to export a CSV report of the results.

.EXAMPLE
    .\Grant-MailboxDelegation.ps1 -Delegate "john.doe@contoso.com" -MailboxListPath ".\mailboxes.txt"

.LINK
    https://github.com/Dalkin69/Grant-MailboxDelegation
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Delegate,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$MailboxListPath,

    [bool]$AutoMapping = $true,

    [string]$ReportPath
)

#region --- Prerequisites ---------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error "The ExchangeOnlineManagement module is not installed. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser"
    exit 1
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect only if no active Exchange Online session exists
$connection = Get-ConnectionInformation -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Connected' }
if (-not $connection) {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    try {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
    catch {
        Write-Error "Unable to connect to Exchange Online: $($_.Exception.Message)"
        exit 1
    }
}

# Check that the delegate exists
try {
    $delegateObj = Get-EXORecipient -Identity $Delegate -ErrorAction Stop
    $delegateUpn = $delegateObj.PrimarySmtpAddress
    Write-Host "Delegate found: $($delegateObj.DisplayName) <$delegateUpn>" -ForegroundColor Cyan
}
catch {
    Write-Error "Delegate '$Delegate' not found: $($_.Exception.Message)"
    exit 1
}

#endregion

#region --- Read input file -------------------------------------------------------

$mailboxes = Get-Content -Path $MailboxListPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') } |
    Sort-Object -Unique

if (-not $mailboxes) {
    Write-Warning "No mailbox found in '$MailboxListPath'."
    exit 0
}

Write-Host "$($mailboxes.Count) mailbox(es) to process.`n" -ForegroundColor Cyan

#endregion

#region --- Processing ------------------------------------------------------------

$results = [System.Collections.Generic.List[object]]::new()

foreach ($mbx in $mailboxes) {

    $result = [PSCustomObject]@{
        Mailbox    = $mbx
        Delegate   = $delegateUpn
        FullAccess = ''
        SendAs     = ''
        Error      = ''
    }

    Write-Host "Processing $mbx" -ForegroundColor White

    # Check the mailbox exists
    try {
        $mailbox = Get-EXOMailbox -Identity $mbx -ErrorAction Stop
    }
    catch {
        $result.FullAccess = 'Skipped'
        $result.SendAs     = 'Skipped'
        $result.Error      = "Mailbox not found: $($_.Exception.Message)"
        Write-Warning "  Mailbox not found, skipped."
        $results.Add($result)
        continue
    }

    # --- Full Access ---
    try {
        $existingFA = Get-EXOMailboxPermission -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop |
            Where-Object { $_.User -eq $delegateUpn -and $_.AccessRights -contains 'FullAccess' -and -not $_.Deny }

        if ($existingFA) {
            $result.FullAccess = 'Already present'
            Write-Host "  FullAccess : already present" -ForegroundColor DarkGray
        }
        elseif ($PSCmdlet.ShouldProcess($mbx, "Grant FullAccess to $delegateUpn")) {
            Add-MailboxPermission -Identity $mailbox.PrimarySmtpAddress -User $delegateUpn `
                -AccessRights FullAccess -InheritanceType All -AutoMapping $AutoMapping -ErrorAction Stop | Out-Null
            $result.FullAccess = 'Granted'
            Write-Host "  FullAccess : granted" -ForegroundColor Green
        }
        else {
            $result.FullAccess = 'WhatIf'
        }
    }
    catch {
        $result.FullAccess = 'Failed'
        $result.Error     += "FullAccess: $($_.Exception.Message) "
        Write-Warning "  FullAccess : failed - $($_.Exception.Message)"
    }

    # --- Send As ---
    try {
        $existingSA = Get-EXORecipientPermission -Identity $mailbox.PrimarySmtpAddress -Trustee $delegateUpn `
            -AccessRights SendAs -ErrorAction SilentlyContinue

        if ($existingSA) {
            $result.SendAs = 'Already present'
            Write-Host "  SendAs     : already present" -ForegroundColor DarkGray
        }
        elseif ($PSCmdlet.ShouldProcess($mbx, "Grant SendAs to $delegateUpn")) {
            Add-RecipientPermission -Identity $mailbox.PrimarySmtpAddress -Trustee $delegateUpn `
                -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
            $result.SendAs = 'Granted'
            Write-Host "  SendAs     : granted" -ForegroundColor Green
        }
        else {
            $result.SendAs = 'WhatIf'
        }
    }
    catch {
        $result.SendAs = 'Failed'
        $result.Error += "SendAs: $($_.Exception.Message)"
        Write-Warning "  SendAs     : failed - $($_.Exception.Message)"
    }

    $results.Add($result)
}

#endregion

#region --- Summary ---------------------------------------------------------------

Write-Host "`n===== Summary =====" -ForegroundColor Cyan
$results | Format-Table Mailbox, FullAccess, SendAs -AutoSize

$failed = @($results | Where-Object { $_.FullAccess -eq 'Failed' -or $_.SendAs -eq 'Failed' -or $_.FullAccess -eq 'Skipped' })
if ($failed.Count -gt 0) {
    Write-Warning "$($failed.Count) mailbox(es) had errors. See the report or the Error column for details."
}

if ($ReportPath) {
    try {
        $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Host "Report exported to $ReportPath" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Unable to export report: $($_.Exception.Message)"
    }
}

Write-Host "Note: permissions may take up to 60 minutes to become effective." -ForegroundColor Yellow

#endregion
