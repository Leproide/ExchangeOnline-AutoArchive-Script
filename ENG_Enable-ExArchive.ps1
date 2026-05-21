#Requires -Version 5.1
<#
.LINK
https://github.com/Leproide/ExchangeOnline-AutoArchive-Script

.NOTES
GPL v3 https://www.gnu.org/licenses/gpl-3.0.html

.SYNOPSIS
    Enables the Exchange Online archive, sets the retention policy and starts archiving for a user.

.DESCRIPTION
    Interactive script that:
    - Checks for the ExchangeOnlineManagement module
    - Logs in using postmaster / admin credentials
    - Asks for the target user
    - Asks for the archiving threshold (1-6 years or custom)
    - Shows a summary and asks for confirmation before proceeding

.NOTES
    Requires: ExchangeOnlineManagement PowerShell module
    Minimum roles required: Mail Recipients, Retention Management
    Documentazione: https://learn.microsoft.com/en-us/purview/enable-archive-mailboxes
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# UTILITY FUNCTIONS
# ---------------------------------------------------------------------------

function Write-Header {
    param([string]$Text)
    $line = "=" * 60
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step {
    param([string]$Text)
    Write-Host "[*] $Text" -ForegroundColor Yellow
}

function Write-OK {
    param([string]$Text)
    Write-Host "[OK] $Text" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Text)
    Write-Host "[ERROR] $Text" -ForegroundColor Red
}

function Write-Info {
    param([string]$Text)
    Write-Host "     $Text" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# STEP 1 — Check ExchangeOnlineManagement module
# ---------------------------------------------------------------------------

Write-Header "Exchange Online - Enable User Archive"

Write-Step "Checking ExchangeOnlineManagement module..."

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1

if (-not $module) {
    Write-Fail "The ExchangeOnlineManagement module is NOT installed."
    Write-Host ""
    Write-Host "  To install it, open PowerShell as Administrator and run:" -ForegroundColor White
    Write-Host "  Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host ""
    exit 1
}

Write-OK "Module found: ExchangeOnlineManagement v$($module.Version)"

# ---------------------------------------------------------------------------
# STEP 2 — Login with postmaster credentials
# ---------------------------------------------------------------------------

Write-Header "Connecting to Exchange Online"

Write-Step "Enter the postmaster / Exchange administrator account credentials."
Write-Info "The Microsoft authentication window will open."
Write-Host ""

try {
    # Connect-ExchangeOnline opens the OAuth2 browser popup if -Credential is not passed
    Connect-ExchangeOnline -ShowBanner:$false
    Write-OK "Connection to Exchange Online established."
} catch {
    Write-Fail "Connection failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Enable Auto-Expanding Archive at tenant level
# ---------------------------------------------------------------------------

Write-Step "Enabling Auto-Expanding Archive at organization level..."
$orgCfgCheck = Get-OrganizationConfig
if ($orgCfgCheck.AutoExpandingArchiveEnabled) {
    Write-Info "Auto-Expanding Archive already enabled at organization level, step skipped."
} else {
    try {
        Set-OrganizationConfig -AutoExpandingArchive -ErrorAction Stop
        Write-OK "Auto-Expanding Archive enabled at organization level."
    } catch {
        Write-Fail "Unable to enable Auto-Expanding Archive at organization level: $_"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
}

Write-Step "Verifying Auto-Expanding Archive at organization level..."
$orgCfg = Get-OrganizationConfig
Write-OK "AutoExpandingArchiveEnabled = $($orgCfg.AutoExpandingArchiveEnabled)"

# ---------------------------------------------------------------------------
# STEP 3 — Target user selection
# ---------------------------------------------------------------------------

Write-Header "User Selection"

do {
    $userInput = Read-Host "Enter the user UPN or display name (e.g. john.doe@contoso.com)"
    $userInput = $userInput.Trim()

    if ([string]::IsNullOrWhiteSpace($userInput)) {
        Write-Fail "The address cannot be empty. Please try again."
        $mailbox = $null
        continue
    }

    Write-Step "Looking up mailbox for '$userInput'..."

    try {
        $mailbox = Get-Mailbox -Identity $userInput -ErrorAction Stop
    } catch {
        Write-Fail "User not found: $_"
        $mailbox = $null
    }

    if ($mailbox) {
        Write-OK "Found: $($mailbox.DisplayName) <$($mailbox.PrimarySmtpAddress)>"
    }

} while (-not $mailbox)

# Check whether the archive is already active
$archiveAlreadyEnabled = ($mailbox.ArchiveStatus -eq "Active")

if ($archiveAlreadyEnabled) {
    Write-Host ""
    Write-Host "  [i] The online archive is ALREADY active for this user." -ForegroundColor Magenta
    Write-Host "      Only the retention policy will be updated and synchronization started." -ForegroundColor Magenta
}

# ---------------------------------------------------------------------------
# STEP 4 — Archiving threshold selection
# ---------------------------------------------------------------------------

Write-Header "Retention Configuration (Archiving Threshold)"

Write-Host "  Select after how many years emails are moved to the archive:" -ForegroundColor White
Write-Host ""
Write-Host "    [1]  1 year   (365 days)" -ForegroundColor White
Write-Host "    [2]  2 years  (730 days)  -- Default MRM Policy" -ForegroundColor White
Write-Host "    [3]  3 years  (1095 days)" -ForegroundColor White
Write-Host "    [4]  4 years  (1460 days)" -ForegroundColor White
Write-Host "    [5]  5 years  (1825 days)" -ForegroundColor White
Write-Host "    [6]  6 years  (2190 days)" -ForegroundColor White
Write-Host "    [7]  Custom (enter the number of years)" -ForegroundColor White
Write-Host ""

do {
    $choice = Read-Host "Choice (1-7)"
    $retentionDays = 0
    switch ($choice.Trim()) {
        "1" { $retentionYears = 1 ; $retentionDays = 365  ; break }
        "2" { $retentionYears = 2 ; $retentionDays = 730  ; break }
        "3" { $retentionYears = 3 ; $retentionDays = 1095 ; break }
        "4" { $retentionYears = 4 ; $retentionDays = 1460 ; break }
        "5" { $retentionYears = 5 ; $retentionDays = 1825 ; break }
        "6" { $retentionYears = 6 ; $retentionDays = 2190 ; break }
        "7" {
            do {
                $customInput = Read-Host "  Enter the number of years (e.g. 7, 10, 15...)"
                if ($customInput -match '^\d+$' -and [int]$customInput -ge 1 -and [int]$customInput -le 99) {
                    $retentionYears = [int]$customInput
                    $retentionDays  = $retentionYears * 365
                } else {
                    Write-Fail "Invalid value. Enter an integer between 1 and 99."
                    $retentionDays = 0
                }
            } while ($retentionDays -eq 0)
            break
        }
        default {
            Write-Fail "Invalid choice. Enter a number from 1 to 7."
            $retentionDays = 0
        }
    }
} while ($retentionDays -eq 0)

Write-OK "Selected threshold: $retentionYears year(s) ($retentionDays days)"

# Tag/policy names — reused if they already exist
$tagName    = "Archive after $retentionYears year"
if ($retentionYears -gt 1) { $tagName = "Archive after $retentionYears years" }
$policyName = "Policy-Archive-$retentionYears-Year"
if ($retentionYears -gt 1) { $policyName = "Policy-Archive-$retentionYears-Years" }

# ---------------------------------------------------------------------------
# STEP 5 — Summary and confirmation
# ---------------------------------------------------------------------------

Write-Header "Operations Summary"

Write-Host "  User                : $($mailbox.DisplayName)" -ForegroundColor White
Write-Host "  UPN                 : $($mailbox.PrimarySmtpAddress)" -ForegroundColor White
Write-Host "  Online archive      : $(if ($archiveAlreadyEnabled) { 'Already active (no changes)' } else { 'TO BE ENABLED' })" -ForegroundColor White
Write-Host "  Retention tag       : '$tagName'" -ForegroundColor White
Write-Host "  Retention policy    : '$policyName'" -ForegroundColor White
Write-Host "  Archiving threshold : emails older than $retentionYears year(s)" -ForegroundColor White
Write-Host "  Action              : MoveToArchive (moves to archive, does not delete)" -ForegroundColor White
Write-Host "  Immediate trigger   : Start-ManagedFolderAssistant (forced)" -ForegroundColor White
Write-Host ""
Write-Host "  NOTE: If the retention tag or policy already exist they will be reused." -ForegroundColor DarkYellow
Write-Host ""

$confirm = Read-Host "Proceed? (Y to confirm / N to cancel)"

if ($confirm.Trim().ToUpper() -ne "Y") {
    Write-Host ""
    Write-Host "  Operation cancelled by user." -ForegroundColor Magenta
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 0
}

# ---------------------------------------------------------------------------
# STEP 6 — Execution
# ---------------------------------------------------------------------------

Write-Header "Execution"

# 6a. Enable archive if not active
if (-not $archiveAlreadyEnabled) {
    Write-Step "Enabling online archive..."
    try {
        Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -Archive -ErrorAction Stop | Out-Null
		Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -ArchiveName "Online Archive - $($mailbox.PrimarySmtpAddress)" -ErrorAction Stop | Out-Null
        Write-OK "Online archive enabled."
    } catch {
        Write-Fail "Unable to enable the archive: $_"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
} else {
    Write-Info "Online archive already active, step skipped."
}

# 6a-bis. Enable Auto-Expanding on the mailbox
Write-Step "Enabling Auto-Expanding Archive on the mailbox..."
try {
    Enable-Mailbox -Identity $mailbox.PrimarySmtpAddress -AutoExpandingArchive -ErrorAction Stop | Out-Null
    Write-OK "Auto-Expanding Archive enabled for the mailbox."
} catch {
    Write-Fail "Unable to enable Auto-Expanding Archive on the mailbox: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6b. Create the retention tag (if it does not exist)
Write-Step "Checking/creating retention tag '$tagName'..."
try {
    $existingTag = Get-RetentionPolicyTag -Identity $tagName -ErrorAction SilentlyContinue
    if ($existingTag) {
        Write-Info "Tag already exists, reusing."
    } else {
        New-RetentionPolicyTag `
            -Name            $tagName `
            -Type            All `
            -RetentionEnabled $true `
            -AgeLimitForRetention $retentionDays `
            -RetentionAction MoveToArchive `
            -ErrorAction Stop | Out-Null
        Write-OK "Retention tag created."
    }
} catch {
    Write-Fail "Error managing the retention tag: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6c. Create the retention policy (if it does not exist) and link the tag
Write-Step "Checking/creating retention policy '$policyName'..."
try {
    $existingPolicy = Get-RetentionPolicy -Identity $policyName -ErrorAction SilentlyContinue
    if ($existingPolicy) {
        Write-Info "Policy already exists, reusing."
        # Ensure the tag is present in the policy
        $tagLinks = $existingPolicy.RetentionPolicyTagLinks
        if ($tagLinks -notcontains $tagName) {
            Set-RetentionPolicy -Identity $policyName -RetentionPolicyTagLinks ($tagLinks + $tagName) -ErrorAction Stop | Out-Null
            Write-Info "Tag added to existing policy."
        }
    } else {
        New-RetentionPolicy `
            -Name                   $policyName `
            -RetentionPolicyTagLinks $tagName `
            -ErrorAction Stop | Out-Null
        Write-OK "Retention policy created."
    }
} catch {
    Write-Fail "Error managing the retention policy: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6d. Assign the policy to the mailbox
Write-Step "Assigning policy '$policyName' to the mailbox..."
try {
    Set-Mailbox -Identity $mailbox.PrimarySmtpAddress -RetentionPolicy $policyName -ErrorAction Stop
    Write-OK "Policy assigned."
} catch {
    Write-Fail "Error assigning the policy: $_"
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

# 6e. Force immediate run of the Managed Folder Assistant
Write-Step "Forcing immediate Managed Folder Assistant run (archiving triggered)..."
try {
    Start-ManagedFolderAssistant -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop
    Write-OK "Managed Folder Assistant started."
    Write-Info "Archiving will begin within a few minutes."
    Write-Info "To monitor progress, check the 'In-Place Archive' folder in OWA."
} catch {
    Write-Fail "Error starting the Managed Folder Assistant: $_"
    Write-Info "The policy has been applied anyway. Archiving will start automatically within 7 days."
}


# ---------------------------------------------------------------------------
# STEP 7 — Final verification and summary
# ---------------------------------------------------------------------------

Write-Header "Final Verification"

Write-Step "Retrieving updated mailbox state..."
try {
    $updatedMailbox = Get-Mailbox -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop

    Write-Host ""
    Write-Host "  --- Final State ---" -ForegroundColor Cyan
    Write-Host "  User             : $($updatedMailbox.DisplayName)" -ForegroundColor White
    Write-Host "  UPN              : $($updatedMailbox.PrimarySmtpAddress)" -ForegroundColor White
    Write-Host "  Archive status   : $($updatedMailbox.ArchiveStatus)" -ForegroundColor White
    Write-Host "  Archive name     : $($updatedMailbox.ArchiveName)" -ForegroundColor White
    Write-Host "  Retention policy : $($updatedMailbox.RetentionPolicy)" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Info "Unable to retrieve final summary: $_"
}

Write-Host "  Operation completed successfully!" -ForegroundColor Green
Write-Host ""

# ---------------------------------------------------------------------------
# Disconnect
# ---------------------------------------------------------------------------

Write-Step "Disconnecting from Exchange Online..."
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
Write-OK "Session closed."
Write-Host ""
