#Requires -Version 5.1
<#
.SYNOPSIS
    Sets up Exchange Online access policy for the OpenClaw M365 Entra app.

.DESCRIPTION
    This script installs the ExchangeOnlineManagement module, connects to Exchange Online,
    and configures an RBAC-for-Applications policy so your registered Entra app can access
    only the mailboxes you explicitly grant it access to.

    Two modes are available:
        1. RBAC for Applications (recommended – current Microsoft standard)
        2. Legacy Application Access Policy (deprecated – will be removed by Microsoft)

    References:
        https://learn.microsoft.com/exchange/permissions-exo/application-rbac

.PARAMETER AppId
    Application (client) ID of the Entra app registration (CLIENT_ID).

.PARAMETER ObjectId
    Object ID of the Entra app's service principal in your tenant.
    Find it under: Entra admin center → Enterprise applications → <your app> → Overview → Object ID

.PARAMETER GroupEmailAddress
    E-mail address of an existing mail-enabled security group whose members the app may access.
    If the group does not yet exist, this script can create it for you.

.PARAMETER PolicyMode
    Choose 'RBAC' (recommended) or 'Legacy'.
    Default: RBAC

.EXAMPLE
    # Interactive – the script prompts for every value
    .\setup-exchange-policy.ps1

.EXAMPLE
    # Non-interactive (all parameters supplied)
    .\setup-exchange-policy.ps1 `
        -AppId       "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
        -ObjectId    "yyyyyyyy-yyyy-yyyy-yyyy-yyyyyyyyyyyy" `
        -GroupEmailAddress "openclaw-m365@yourdomain.com"
#>

[CmdletBinding()]
param(
    [string]$AppId,
    [string]$ObjectId,
    [string]$GroupEmailAddress,
    [ValidateSet("RBAC", "Legacy")]
    [string]$PolicyMode = "RBAC"
)

# ---------------------------------------------------------------------------
# Helper: prompt for a value if not already provided
# ---------------------------------------------------------------------------
function Read-RequiredParam {
    param(
        [string]$Value,
        [string]$Prompt,
        [string]$Default = ""
    )
    if ($Value -ne "") { return $Value }
    $display = if ($Default -ne "") { "$Prompt [default: $Default]" } else { $Prompt }
    $input = Read-Host $display
    if ($input -eq "" -and $Default -ne "") { return $Default }
    while ($input -eq "") {
        Write-Warning "This value is required."
        $input = Read-Host $display
    }
    return $input
}

# ---------------------------------------------------------------------------
# Banner
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  OpenClaw M365 – Exchange Online Access Policy Setup" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script configures Exchange Online so your Entra app can"
Write-Host "access only the mailboxes inside a specific security group."
Write-Host ""

# ---------------------------------------------------------------------------
# Step 0 – Install / update ExchangeOnlineManagement module
# ---------------------------------------------------------------------------
Write-Host "[Step 1/6] Checking ExchangeOnlineManagement module ..." -ForegroundColor Yellow

$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version | Select-Object -Last 1
if (-not $module) {
    Write-Host "  Module not found. Installing..." -ForegroundColor Gray
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    Write-Host "  Module installed." -ForegroundColor Green
} else {
    Write-Host "  Found version $($module.Version). Checking for updates..." -ForegroundColor Gray
    try {
        Update-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "  Could not update the module – using installed version $($module.Version)." -ForegroundColor Gray
    }
}
Import-Module ExchangeOnlineManagement -ErrorAction Stop
Write-Host "  ExchangeOnlineManagement loaded." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 1 – Collect parameters
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[Step 2/6] Collecting required parameters..." -ForegroundColor Yellow
Write-Host ""

$AppId             = Read-RequiredParam -Value $AppId             -Prompt "  Application (client) ID (CLIENT_ID from Entra portal)"
$ObjectId          = Read-RequiredParam -Value $ObjectId          -Prompt "  Service principal Object ID (Entra → Enterprise apps → <app> → Object ID)"
$GroupEmailAddress = Read-RequiredParam -Value $GroupEmailAddress -Prompt "  E-mail address of the mail-enabled security group (will be created if missing)"

if ($PolicyMode -eq "RBAC") {
    Write-Host ""
    Write-Host "  Mode: RBAC for Applications (recommended)" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "  Mode: Legacy Application Access Policy (deprecated)" -ForegroundColor Yellow
    Write-Host "  WARNING: Microsoft is retiring Application Access Policies." -ForegroundColor Yellow
    Write-Host "           Migrate to RBAC for Applications as soon as possible." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  Summary of values to be used:"
Write-Host "    AppId             : $AppId"
Write-Host "    ObjectId          : $ObjectId"
Write-Host "    Group email       : $GroupEmailAddress"
Write-Host "    Policy mode       : $PolicyMode"
Write-Host ""

$confirm = Read-Host "  Proceed? [Y/n]"
if ($confirm -ne "" -and $confirm -notmatch "^[Yy]") {
    Write-Host "Aborted." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Step 2 – Connect to Exchange Online
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[Step 3/6] Connecting to Exchange Online..." -ForegroundColor Yellow
Write-Host "  Using Device Code flow (works on Mac and headless Linux)."
Write-Host "  A code will appear – open https://microsoft.com/devicelogin in a browser"
Write-Host "  and sign in with a Global Administrator or Exchange Administrator account."
Write-Host ""

Connect-ExchangeOnline -Device -ShowBanner:$false

Write-Host "  Connected to Exchange Online." -ForegroundColor Green

# ---------------------------------------------------------------------------
# Step 3 – Create mail-enabled security group (if not existing)
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[Step 4/6] Verifying security group '$GroupEmailAddress'..." -ForegroundColor Yellow

$group = Get-DistributionGroup -Identity $GroupEmailAddress -ErrorAction SilentlyContinue
if (-not $group) {
    Write-Host "  Group not found. Creating a new mail-enabled security group..." -ForegroundColor Gray

    $groupAlias = ($GroupEmailAddress -split "@")[0] -replace "[^a-zA-Z0-9-]", "-"
    $groupName  = Read-Host "  Display name for the new group [default: OpenClaw M365 Exchange Access]"
    if ($groupName -eq "") { $groupName = "OpenClaw M365 Exchange Access" }

    New-DistributionGroup `
        -Name  $groupName `
        -Alias $groupAlias `
        -PrimarySmtpAddress $GroupEmailAddress `
        -Type  Security

    Write-Host "  Group '$groupName' created." -ForegroundColor Green

    $addMembers = Read-Host "  Add mailbox members to the group now? [Y/n]"
    if ($addMembers -eq "" -or $addMembers -match "^[Yy]") {
        $more = "Y"
        while ($more -match "^[Yy]") {
            $member = Read-Host "    Enter member e-mail address (UPN)"
            if ($member -ne "") {
                Add-DistributionGroupMember -Identity $GroupEmailAddress -Member $member
                Write-Host "    Added: $member" -ForegroundColor Green
            }
            $more = Read-Host "    Add another member? [y/N]"
        }
    }
} else {
    Write-Host "  Group found: $($group.DisplayName) ($($group.PrimarySmtpAddress))" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Step 4 – Create access policy
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[Step 5/6] Creating access policy ($PolicyMode mode)..." -ForegroundColor Yellow

if ($PolicyMode -eq "RBAC") {

    # ---- RBAC for Applications (recommended) --------------------------------

    # 4a. Ensure service principal exists in Exchange
    Write-Host "  Checking / creating Exchange service principal..." -ForegroundColor Gray
    $sp = Get-ServicePrincipal -Identity $AppId -ErrorAction SilentlyContinue
    if (-not $sp) {
        $spDisplayName = Read-Host "  Display name for the service principal [default: OpenClaw M365 App]"
        if ($spDisplayName -eq "") { $spDisplayName = "OpenClaw M365 App" }
        New-ServicePrincipal -AppId $AppId -ObjectId $ObjectId -DisplayName $spDisplayName
        Write-Host "  Service principal created." -ForegroundColor Green
    } else {
        Write-Host "  Service principal already exists: $($sp.DisplayName)" -ForegroundColor Green
    }

    # 4b. Create a management scope limited to the group members
    $scopeName = "OpenClaw-Scope-" + (($GroupEmailAddress -split "@")[0] -replace "[^a-zA-Z0-9-]", "-")
    Write-Host "  Creating management scope '$scopeName'..." -ForegroundColor Gray
    $existingScope = Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue
    if (-not $existingScope) {
        New-ManagementScope `
            -Name $scopeName `
            -RecipientRestrictionFilter "MemberOfGroup -eq '$GroupEmailAddress'"
        Write-Host "  Management scope created." -ForegroundColor Green
    } else {
        Write-Host "  Management scope already exists – reusing." -ForegroundColor Green
    }

    # 4c. Assign the Application Mail.Read role within the scope
    $assignmentName = "OpenClaw-MailRead-" + $AppId.Substring(0, [Math]::Min(8, $AppId.Length))
    Write-Host "  Creating role assignment '$assignmentName'..." -ForegroundColor Gray
    $existingAssignment = Get-ManagementRoleAssignment -Identity $assignmentName -ErrorAction SilentlyContinue
    if (-not $existingAssignment) {
        New-ManagementRoleAssignment `
            -Name $assignmentName `
            -App  $AppId `
            -Role "Application Mail.Read" `
            -CustomResourceScope $scopeName
        Write-Host "  Role assignment created." -ForegroundColor Green
    } else {
        Write-Host "  Role assignment already exists – skipping." -ForegroundColor Green
    }

} else {

    # ---- Legacy Application Access Policy (deprecated) ----------------------

    Write-Host "  Creating legacy Application Access Policy..." -ForegroundColor Gray
    New-ApplicationAccessPolicy `
        -AppId            $AppId `
        -PolicyScopeGroupId $GroupEmailAddress `
        -AccessRight      RestrictAccess `
        -Description      "OpenClaw App – access restricted to group members"

    Write-Host "  Legacy Application Access Policy created." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Step 5 – Verify
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "[Step 6/6] Verifying the configuration..." -ForegroundColor Yellow

if ($PolicyMode -eq "RBAC") {

    Write-Host ""
    Write-Host "  Current management role assignments for the app:" -ForegroundColor Gray
    Get-ManagementRoleAssignment -App $AppId | Format-Table Name, Role, CustomResourceScope -AutoSize

} else {

    Write-Host ""
    Write-Host "  All Application Access Policies:" -ForegroundColor Gray
    Get-ApplicationAccessPolicy | Format-Table AppId, ScopeName, AccessRight -AutoSize

    Write-Host ""
    Write-Host "  Testing access for group e-mail '$GroupEmailAddress':" -ForegroundColor Gray
    $testResult = Test-ApplicationAccessPolicy -Identity $GroupEmailAddress -AppId $AppId
    if ($testResult.AccessCheckResult -eq "Granted") {
        Write-Host "  AccessCheckResult: Granted ✔" -ForegroundColor Green
    } else {
        Write-Host "  AccessCheckResult: $($testResult.AccessCheckResult)" -ForegroundColor Yellow
        Write-Host "  Note: It may take up to 1 hour for the policy to propagate." -ForegroundColor Gray
    }
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Setup complete!" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Your Entra app ($AppId)"
Write-Host "can now access the mailboxes that are members of:"
Write-Host "  $GroupEmailAddress"
Write-Host ""
Write-Host "Note: Policy propagation may take up to 60 minutes in large tenants."
Write-Host ""

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Disconnected from Exchange Online." -ForegroundColor Gray
