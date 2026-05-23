#============================================================================== 
# M365 Assessment + Report Generator (Consultant Version) - FINAL
# Generated for: Khaing Linn Htun
# Date: 2026-05-23
#
# Key Outputs (in timestamped Desktop folder):
#  - 00_Scorecard.csv
#  - 00_Executive_Summary.md
#  - 00_Executive_Summary.html
#  - 40_Roadmap_30_60_90.csv
#  - 41_Findings_RiskRegister.csv
#  - Inventory exports (Users/Licenses/Mailbox/OneDrive/Groups/SecureScore etc.)
#
# Notes:
#  - Read-only collection.
#  - SignInActivity may require AuditLog.Read.All + proper tenant/license/role.
#  - Exchange Online shared mailbox export runs in a separate PowerShell process by default.
#==============================================================================

[CmdletBinding()]
param(
    [ValidateSet('D7','D30','D90','D180')]
    [string]$ReportPeriod = 'D180',

    [double]$MailboxThresholdGB = 2,

    [bool]$IncludeSignInActivity = $true,
    [bool]$IncludeMfaRegistrationReport = $true,

    [bool]$IncludeSharedMailboxes = $true,
    [bool]$RunEXOInSeparateProcess = $true,
    [bool]$DisableWAMForEXO = $true
)

#region ── Helpers ────────────────────────────────────────────────────────────
function Write-Section {
    param([string]$Title)
    Write-Host "" 
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host (" " + $Title) -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
}
function Write-Ok   { param([string]$Msg) Write-Host "[OK]   $Msg" -ForegroundColor Green }
function Write-Warn { param([string]$Msg) Write-Host "[WARN] $Msg" -ForegroundColor Yellow }
function Write-Err  { param([string]$Msg) Write-Host "[ERR]  $Msg" -ForegroundColor Red }

function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Warn "Installing module: $Name"
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Ok "Installed module: $Name"
    } else {
        Write-Ok "Module available: $Name"
    }
}

function Export-JsonSafe {
    param(
        [Parameter(Mandatory=$true)][object]$Object,
        [Parameter(Mandatory=$true)][string]$Path
    )
    try {
        $Object | ConvertTo-Json -Depth 15 | Out-File -FilePath $Path -Encoding UTF8
    } catch {
        "JSON export failed: $($_)" | Out-File -FilePath $Path -Encoding UTF8
    }
}

function Add-Finding {
    param(
        [string]$Category,
        [string]$Finding,
        [ValidateSet('Critical','High','Medium','Low','Info')]$Severity = 'Medium',
        [string]$Evidence = 'Auto-detected',
        [string]$Recommendation = 'Review and apply Microsoft best practice',
        [ValidateSet('Quick','Planned','Project')]$Effort = 'Planned'
    )
    $script:Findings += [PSCustomObject]@{
        Category       = $Category
        Finding        = $Finding
        Severity       = $Severity
        Evidence       = $Evidence
        Recommendation = $Recommendation
        Effort         = $Effort
    }
}

function Severity-Rank {
    param([string]$Severity)
    switch ($Severity) {
        'Critical' { 1 }
        'High'     { 2 }
        'Medium'   { 3 }
        'Low'      { 4 }
        default    { 5 }
    }
}

function Invoke-GraphPaged {
    param(
        [Parameter(Mandatory=$true)][string]$Uri
    )
    $all = @()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($resp.value) { $all += $resp.value }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}
#endregion

#region ── Banner & Output Folder ─────────────────────────────────────────────
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host " Microsoft 365 Assessment + Report Generator (Consultant Version)" -ForegroundColor Cyan
Write-Host (" " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$outputRoot = Join-Path $env:USERPROFILE "Desktop\M365_Assessment_Consultant_$timestamp"
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
Write-Ok "Output folder: $outputRoot"
#endregion

#region ── Modules ───────────────────────────────────────────────────────────
Write-Section "Prerequisites - Modules"
$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Reports',
    'Microsoft.Graph.Security'
)
foreach ($m in $graphModules) { Ensure-Module $m }
#endregion

#region ── Graph Connect ─────────────────────────────────────────────────────
Write-Section "Connecting - Microsoft Graph"
$scopes = @(
    'User.Read.All',
    'Directory.Read.All',
    'Reports.Read.All',
    'Group.Read.All',
    'Organization.Read.All',
    'SecurityEvents.Read.All',
    'Policy.Read.All'
)
if ($IncludeSignInActivity) { $scopes += 'AuditLog.Read.All' }

try {
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    Write-Ok "Connected to Microsoft Graph"
} catch {
    Write-Err "Failed to connect to Microsoft Graph: $($_)"
    throw
}
#endregion

#region ── Initialize collections ─────────────────────────────────────────────
$Findings = @()
$Scorecard = [ordered]@{}
$Roadmap = @()
#endregion

#==============================================================================
# A1. Users Export + Guest Summary
#==============================================================================
Write-Section "A1. Users Export + Guest Summary"
$usersExport = $null
$rawUsers = $null
$guestCount = 0

$baseUserProps = @(
    'Id','DisplayName','UserPrincipalName','Mail','GivenName','Surname',
    'JobTitle','Department','OfficeLocation','CompanyName',
    'City','State','Country','PostalCode','StreetAddress',
    'MobilePhone','BusinessPhones','AccountEnabled','UserType',
    'CreatedDateTime','AssignedLicenses','AssignedPlans',
    'OnPremisesSyncEnabled','OnPremisesLastSyncDateTime',
    'OnPremisesDomainName','OnPremisesSamAccountName',
    'ProxyAddresses','OtherMails','UsageLocation','PreferredLanguage',
    'EmployeeId','EmployeeType'
)

try {
    $props = @($baseUserProps)
    if ($IncludeSignInActivity) { $props += 'SignInActivity' }

    $rawUsers = Get-MgUser -All -Property $props -ConsistencyLevel eventual

    $selectProps = @(
        'Id','DisplayName','UserPrincipalName','Mail','GivenName','Surname',
        'JobTitle','Department','OfficeLocation','CompanyName',
        'City','State','Country','PostalCode','StreetAddress',
        'MobilePhone',
        @{Name='BusinessPhones';Expression={$_.BusinessPhones -join '; '}},
        'AccountEnabled','UserType','CreatedDateTime',
        @{Name='LicenseCount';Expression={($_.AssignedLicenses).Count}},
        @{Name='Licensed';Expression={ if(($_.AssignedLicenses).Count -gt 0){'Yes'} else {'No'} }},
        'OnPremisesSyncEnabled','OnPremisesLastSyncDateTime','OnPremisesDomainName','OnPremisesSamAccountName',
        @{Name='ProxyAddresses';Expression={$_.ProxyAddresses -join '; '}},
        @{Name='OtherMails';Expression={$_.OtherMails -join '; '}},
        'UsageLocation','PreferredLanguage','EmployeeId','EmployeeType'
    )

    if ($IncludeSignInActivity) {
        $selectProps += @{Name='LastSignInDateTime';Expression={$_.SignInActivity.LastSignInDateTime}}
        $selectProps += @{Name='LastNonInteractiveSignIn';Expression={$_.SignInActivity.LastNonInteractiveSignInDateTime}}
    } else {
        $selectProps += @{Name='LastSignInDateTime';Expression={$null}}
        $selectProps += @{Name='LastNonInteractiveSignIn';Expression={$null}}
    }

    $usersExport = $rawUsers | Select-Object -Property $selectProps
    $usersExport | Export-Csv -Path (Join-Path $outputRoot '01_All_Users.csv') -NoTypeInformation -Encoding UTF8

    $guestCount = ($rawUsers | Where-Object { $_.UserType -eq 'Guest' }).Count

    $Scorecard['Total Users'] = $usersExport.Count
    $Scorecard['Guest Users'] = $guestCount

    Write-Ok "Users exported: $($usersExport.Count)"
    Write-Ok "Guest users: $guestCount"
}
catch {
    Write-Warn "Users export failed (often SignInActivity permissions). Retrying without SignInActivity..."
    try {
        $IncludeSignInActivity = $false
        $rawUsers = Get-MgUser -All -Property $baseUserProps -ConsistencyLevel eventual

        $selectProps2 = @(
            'Id','DisplayName','UserPrincipalName','Mail','GivenName','Surname',
            'JobTitle','Department','OfficeLocation','CompanyName',
            'City','State','Country','PostalCode','StreetAddress',
            'MobilePhone',
            @{Name='BusinessPhones';Expression={$_.BusinessPhones -join '; '}},
            'AccountEnabled','UserType','CreatedDateTime',
            @{Name='LicenseCount';Expression={($_.AssignedLicenses).Count}},
            @{Name='Licensed';Expression={ if(($_.AssignedLicenses).Count -gt 0){'Yes'} else {'No'} }},
            'OnPremisesSyncEnabled','OnPremisesLastSyncDateTime','OnPremisesDomainName','OnPremisesSamAccountName',
            @{Name='ProxyAddresses';Expression={$_.ProxyAddresses -join '; '}},
            @{Name='OtherMails';Expression={$_.OtherMails -join '; '}},
            'UsageLocation','PreferredLanguage','EmployeeId','EmployeeType',
            @{Name='LastSignInDateTime';Expression={$null}},
            @{Name='LastNonInteractiveSignIn';Expression={$null}}
        )

        $usersExport = $rawUsers | Select-Object -Property $selectProps2
        $usersExport | Export-Csv -Path (Join-Path $outputRoot '01_All_Users.csv') -NoTypeInformation -Encoding UTF8

        $guestCount = ($rawUsers | Where-Object { $_.UserType -eq 'Guest' }).Count
        $Scorecard['Total Users'] = $usersExport.Count
        $Scorecard['Guest Users'] = $guestCount

        Write-Ok "Users exported without SignInActivity: $($usersExport.Count)"
    } catch {
        Write-Err "Users export failed: $($_)"
        Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '01_Users_Error.json')
    }
}

#==============================================================================
# A2. Mailbox Usage
#==============================================================================
Write-Section "A2. Mailbox Usage (Reports API)"
$mailbox = $null
try {
    $mailboxUri = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$ReportPeriod')"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $mailboxUri -OutputType HttpResponseMessage
    $csv  = $resp.Content.ReadAsStringAsync().Result
    $raw  = $csv | ConvertFrom-Csv

    $mailbox = $raw | Select-Object -Property @(
        @{Name='UserPrincipalName';Expression={$_.'User Principal Name'}},
        @{Name='DisplayName';Expression={$_.'Display Name'}},
        @{Name='IsDeleted';Expression={$_.'Is Deleted'}},
        @{Name='LastActivityDate';Expression={$_.'Last Activity Date'}},
        @{Name='ItemCount';Expression={$_.'Item Count'}},
        @{Name='StorageUsed_Bytes';Expression={$_.'Storage Used (Byte)'}},
        @{Name='StorageUsed_GB';Expression={[math]::Round([double]$_.'Storage Used (Byte)'/1GB,2)}},
        @{Name='HasArchive';Expression={$_.'Has Archive'}},
        @{Name='ReportPeriod';Expression={$_.'Report Period'}}
    )

    $mailbox | Export-Csv -Path (Join-Path $outputRoot '02_Mailbox_Usage.csv') -NoTypeInformation -Encoding UTF8
    $Scorecard['Mailboxes Reported'] = $mailbox.Count
    Write-Ok "Mailbox usage rows: $($mailbox.Count)"
}
catch {
    Write-Err "Mailbox usage export failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '02_MailboxUsage_Error.json')
}

#==============================================================================
# A3. OneDrive Usage
#==============================================================================
Write-Section "A3. OneDrive Usage (Reports API)"
$od = $null
try {
    $odUri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$ReportPeriod')"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $odUri -OutputType HttpResponseMessage
    $csv  = $resp.Content.ReadAsStringAsync().Result
    $raw  = $csv | ConvertFrom-Csv

    $od = $raw | Select-Object -Property @(
        @{Name='OwnerPrincipalName';Expression={$_.'Owner Principal Name'}},
        @{Name='OwnerDisplayName';Expression={$_.'Owner Display Name'}},
        @{Name='IsDeleted';Expression={$_.'Is Deleted'}},
        @{Name='LastActivityDate';Expression={$_.'Last Activity Date'}},
        @{Name='FileCount';Expression={$_.'File Count'}},
        @{Name='StorageUsed_Bytes';Expression={$_.'Storage Used (Byte)'}},
        @{Name='StorageUsed_GB';Expression={[math]::Round([double]$_.'Storage Used (Byte)'/1GB,2)}},
        @{Name='SiteURL';Expression={$_.'Site URL'}},
        @{Name='ReportPeriod';Expression={$_.'Report Period'}}
    )

    $od | Export-Csv -Path (Join-Path $outputRoot '03_OneDrive_Usage.csv') -NoTypeInformation -Encoding UTF8
    $Scorecard['OneDrive Accounts Reported'] = $od.Count
    Write-Ok "OneDrive usage rows: $($od.Count)"
}
catch {
    Write-Err "OneDrive usage export failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '03_OneDriveUsage_Error.json')
}

#==============================================================================
# A4. Mailboxes Under Threshold
#==============================================================================
Write-Section "A4. Mailboxes Under Threshold"
try {
    if ($mailbox) {
        $under = $mailbox | Where-Object { [double]$_.StorageUsed_GB -lt $MailboxThresholdGB } | Sort-Object StorageUsed_GB
        $under | Export-Csv -Path (Join-Path $outputRoot '04_Mailbox_Under_Threshold.csv') -NoTypeInformation -Encoding UTF8
        $Scorecard["Mailboxes Under ${MailboxThresholdGB} GB"] = $under.Count
        Write-Ok "Mailboxes under $MailboxThresholdGB GB: $($under.Count)"
    }
}
catch {
    Write-Err "Threshold filter failed: $($_)"
}

#==============================================================================
# A5. License Inventory
#==============================================================================
Write-Section "A5. Product / License Inventory"
try {
    $SkuFriendly = @{
        'O365_BUSINESS_ESSENTIALS'          = 'Microsoft 365 Business Basic'
        'O365_BUSINESS_PREMIUM'             = 'Microsoft 365 Business Standard'
        'SMB_BUSINESS_PREMIUM'              = 'Microsoft 365 Business Premium'
        'STANDARDPACK'                      = 'Office 365 E1'
        'ENTERPRISEPACK'                    = 'Office 365 E3'
        'ENTERPRISEPREMIUM'                 = 'Office 365 E5'
        'SPE_E3'                            = 'Microsoft 365 E3'
        'SPE_E5'                            = 'Microsoft 365 E5'
        'EXCHANGESTANDARD'                  = 'Exchange Online Plan 1'
        'EXCHANGEENTERPRISE'                = 'Exchange Online Plan 2'
        'AAD_PREMIUM'                       = 'Microsoft Entra ID P1'
        'AAD_PREMIUM_P2'                    = 'Microsoft Entra ID P2'
        'WIN_DEF_ATP'                       = 'Microsoft Defender for Endpoint P2'
        'ATP_ENTERPRISE'                    = 'Microsoft Defender for Office 365 P1'
        'THREAT_INTELLIGENCE'               = 'Microsoft Defender for Office 365 P2'
        'INTUNE_A'                          = 'Microsoft Intune'
        'EMS_E3'                            = 'Enterprise Mobility + Security E3'
        'EMS_E5'                            = 'Enterprise Mobility + Security E5'
    }

    $skus = Get-MgSubscribedSku -All
    $license = $skus | Select-Object -Property @(
        @{Name='SkuId';Expression={$_.SkuId}},
        @{Name='SkuPartNumber';Expression={$_.SkuPartNumber}},
        @{Name='FriendlyName';Expression={ if($SkuFriendly.ContainsKey($_.SkuPartNumber)){$SkuFriendly[$_.SkuPartNumber]} else {$_.SkuPartNumber} }},
        @{Name='AppliesTo';Expression={$_.AppliesTo}},
        @{Name='CapabilityStatus';Expression={$_.CapabilityStatus}},
        @{Name='TotalLicenses';Expression={$_.PrepaidUnits.Enabled}},
        @{Name='ConsumedLicenses';Expression={$_.ConsumedUnits}},
        @{Name='AvailableLicenses';Expression={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}},
        @{Name='ServicePlans';Expression={($_.ServicePlans | Select-Object -ExpandProperty ServicePlanName) -join '; '}}
    )

    $license | Export-Csv -Path (Join-Path $outputRoot '05_License_Inventory.csv') -NoTypeInformation -Encoding UTF8
    $Scorecard['Total SKUs / Products'] = $license.Count
    Write-Ok "SKUs exported: $($license.Count)"
}
catch {
    Write-Err "License inventory failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '05_LicenseInventory_Error.json')
}

#==============================================================================
# A6. Groups
#==============================================================================
Write-Section "A6. Groups (All Types)"
try {
    $groupProps = @(
        'Id','DisplayName','Description','Mail','MailEnabled','MailNickname','SecurityEnabled',
        'GroupTypes','Visibility','MembershipRule','MembershipRuleProcessingState',
        'CreatedDateTime','RenewedDateTime','OnPremisesSyncEnabled','ProxyAddresses'
    )

    $groups = Get-MgGroup -All -Property $groupProps

    $groupSelect = @(
        'Id','DisplayName','Description','Mail','MailEnabled','MailNickname','SecurityEnabled',
        @{Name='GroupTypes';Expression={$_.GroupTypes -join '; '}},
        @{Name='GroupCategory';Expression={
            $gt = $_.GroupTypes
            if ($gt -contains 'Unified') {
                if ($gt -contains 'DynamicMembership') { 'Microsoft 365 Group (Dynamic)' } else { 'Microsoft 365 Group' }
            } elseif ($_.MailEnabled -and $_.SecurityEnabled) {
                'Mail-Enabled Security Group'
            } elseif ($_.MailEnabled -and -not $_.SecurityEnabled) {
                'Distribution Group'
            } elseif (-not $_.MailEnabled -and $_.SecurityEnabled) {
                if ($gt -contains 'DynamicMembership') { 'Dynamic Security Group' } else { 'Security Group' }
            } else {
                'Other'
            }
        }},
        'Visibility','MembershipRule','MembershipRuleProcessingState','CreatedDateTime','RenewedDateTime','OnPremisesSyncEnabled',
        @{Name='ProxyAddresses';Expression={$_.ProxyAddresses -join '; '}}
    )

    $g = $groups | Select-Object -Property $groupSelect
    $g | Export-Csv -Path (Join-Path $outputRoot '09_All_Groups.csv') -NoTypeInformation -Encoding UTF8

    $gSummary = $g | Group-Object GroupCategory | Select-Object @{Name='GroupType';Expression={$_.Name}},Count | Sort-Object Count -Descending
    $gSummary | Export-Csv -Path (Join-Path $outputRoot '09b_Group_Summary_ByType.csv') -NoTypeInformation -Encoding UTF8

    $Scorecard['Total Groups'] = $g.Count
    Write-Ok "Groups exported: $($g.Count)"
}
catch {
    Write-Err "Group export failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '09_Groups_Error.json')
}

#==============================================================================
# B1. Secure Score
#==============================================================================
Write-Section "B1. Microsoft Secure Score (Summary + Controls)"
$controls = $null
$profiles = $null
$secureScorePct = $null
try {
    $scores = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScores?$top=1'
    if ($scores.value -and $scores.value.Count -gt 0) {
        $latest = $scores.value[0]
        $secureScorePct = if([double]$latest.maxScore -gt 0){ [math]::Round(([double]$latest.currentScore/[double]$latest.maxScore)*100,2) } else { 0 }

        [PSCustomObject]@{
            ReportDate      = $latest.createdDateTime
            CurrentScore    = $latest.currentScore
            MaxScore        = $latest.maxScore
            ScorePercentage = "$secureScorePct%"
            EnabledServices = ($latest.enabledServices -join '; ')
        } | Export-Csv -Path (Join-Path $outputRoot '06a_Security_Score_Summary.csv') -NoTypeInformation -Encoding UTF8

        $controls = $latest.controlScores | Select-Object -Property @(
            @{Name='ControlCategory';Expression={$_.controlCategory}},
            @{Name='ControlName';Expression={$_.controlName}},
            @{Name='Score';Expression={$_.score}},
            @{Name='ScoreInPercentage';Expression={$_.scoreInPercentage}},
            @{Name='ImplementationStatus';Expression={$_.implementationStatus}},
            @{Name='Description';Expression={$_.description}},
            @{Name='LastSynced';Expression={$_.lastSynced}}
        )
        $controls | Export-Csv -Path (Join-Path $outputRoot '06b_Security_Score_Controls.csv') -NoTypeInformation -Encoding UTF8

        $Scorecard['Secure Score %'] = $secureScorePct
        Write-Ok "Secure Score: $secureScorePct%"

        if ($secureScorePct -lt 50) {
            Add-Finding -Category 'Security Posture' -Finding 'Secure Score below 50%' -Severity 'High' -Evidence "Secure Score = $secureScorePct%" -Recommendation 'Prioritize Secure Score actions starting with Identity and Email/Data controls; track weekly progress.' -Effort 'Planned'
        }

        $idControls = $controls | Where-Object { $_.ControlCategory -eq 'Identity' }
        if ($idControls) {
            $idAvgPct = [math]::Round((($idControls | Where-Object { $_.ScoreInPercentage -ne $null -and $_.ScoreInPercentage -ne '' } | Measure-Object ScoreInPercentage -Average).Average),2)
            $Scorecard['Identity Controls Avg %'] = $idAvgPct
        }
    }
}
catch {
    Write-Err "Secure Score retrieval failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '06_SecureScore_Error.json')
}

#==============================================================================
# B2. Secure Score Control Profiles
#==============================================================================
Write-Section "B2. Secure Score Control Profiles (Metadata)"
try {
    $p = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/security/secureScoreControlProfiles'
    if ($p.value) {
        $profiles = $p.value | Select-Object -Property @(
            @{Name='Id';Expression={$_.id}},
            @{Name='Title';Expression={$_.title}},
            @{Name='ControlCategory';Expression={$_.controlCategory}},
            @{Name='ActionType';Expression={$_.actionType}},
            @{Name='Service';Expression={$_.service}},
            @{Name='Tier';Expression={$_.tier}},
            @{Name='UserImpact';Expression={$_.userImpact}},
            @{Name='ImplementationCost';Expression={$_.implementationCost}},
            @{Name='Rank';Expression={$_.rank}},
            @{Name='Threats';Expression={$_.threats -join '; '}},
            @{Name='Deprecated';Expression={$_.deprecated}},
            @{Name='Remediation';Expression={$_.remediation}},
            @{Name='RemediationImpact';Expression={$_.remediationImpact}},
            @{Name='ActionUrl';Expression={$_.actionUrl}},
            @{Name='MaxScore';Expression={$_.maxScore}}
        )
        $profiles | Export-Csv -Path (Join-Path $outputRoot '12b_Security_Control_Profiles.csv') -NoTypeInformation -Encoding UTF8
        $Scorecard['Control Profiles Exported'] = $profiles.Count
        Write-Ok "Control profiles exported: $($profiles.Count)"

        $top = $profiles | Where-Object { $_.Deprecated -ne $true } | Sort-Object Rank | Select-Object -First 25
        $top | Export-Csv -Path (Join-Path $outputRoot '30_Top_SecureScore_Recommendations.csv') -NoTypeInformation -Encoding UTF8
        Write-Ok "Top Secure Score recommendations exported: $($top.Count)"
    }
}
catch {
    Write-Err "Control profile export failed: $($_)"
    Export-JsonSafe -Object $_ -Path (Join-Path $outputRoot '12_ControlProfiles_Error.json')
}

#==============================================================================
# C1. Privileged Roles (Admin Assignments)
#==============================================================================
Write-Section "C1. Privileged Roles (Admin Assignments)"
try {
    $defs = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$select=id,displayName"
    $defsMap = @{}
    foreach ($d in $defs) { $defsMap[$d.id] = $d.displayName }

    $assign = Invoke-GraphPaged -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$select=principalId,roleDefinitionId"

    $roleAssignments = @()
    foreach ($a in $assign) {
        $roleName = if ($defsMap.ContainsKey($a.roleDefinitionId)) { $defsMap[$a.roleDefinitionId] } else { $a.roleDefinitionId }
        $roleAssignments += [PSCustomObject]@{
            PrincipalId      = $a.principalId
            RoleDefinitionId = $a.roleDefinitionId
            RoleName         = $roleName
        }
    }

    $roleAssignments | Export-Csv -Path (Join-Path $outputRoot '20_Role_Assignments.csv') -NoTypeInformation -Encoding UTF8

    $globalAdmins = ($roleAssignments | Where-Object { $_.RoleName -eq 'Global Administrator' }).Count
    $Scorecard['Global Admin Assignments'] = $globalAdmins
    Write-Ok "Global Administrator assignments: $globalAdmins"

    if ($globalAdmins -gt 4) {
        Add-Finding -Category 'Identity' -Finding 'Too many Global Administrators' -Severity 'High' -Evidence "Global Administrator assignments = $globalAdmins" -Recommendation 'Reduce Global Admin count to 2-4, create break-glass accounts, and use least privilege / PIM where available.' -Effort 'Planned'
    }
}
catch {
    Write-Warn "Could not retrieve role assignments (permissions vary): $($_)"
    Add-Finding -Category 'Identity' -Finding 'Privileged role inventory not collected' -Severity 'Low' -Evidence "$($_)" -Recommendation 'Grant appropriate read permissions (Directory.Read.All / RoleManagement.Read.Directory) and re-run.' -Effort 'Quick'
}

#==============================================================================
# C2. Conditional Access policies
#==============================================================================
Write-Section "C2. Conditional Access (Policy Inventory)"
try {
    $cap = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
    $policies = $cap.value
    $policies | ConvertTo-Json -Depth 10 | Out-File -FilePath (Join-Path $outputRoot '21_ConditionalAccess_Policies.json') -Encoding UTF8

    $policyCount = if ($policies) { $policies.Count } else { 0 }
    $Scorecard['Conditional Access Policies'] = $policyCount
    Write-Ok "Conditional Access policies: $policyCount"

    $hasMfaPolicy = $false
    foreach ($p in $policies) {
        if ($p.grantControls -and $p.grantControls.builtInControls) {
            if ($p.grantControls.builtInControls -contains 'mfa') { $hasMfaPolicy = $true }
        }
    }

    if ($policyCount -eq 0) {
        Add-Finding -Category 'Identity' -Finding 'No Conditional Access policies found' -Severity 'High' -Evidence 'conditionalAccess/policies returned 0' -Recommendation 'Deploy baseline Conditional Access: MFA for admins, block legacy auth, require strong auth for high-risk sign-ins.' -Effort 'Project'
    } elseif (-not $hasMfaPolicy) {
        Add-Finding -Category 'Identity' -Finding 'Conditional Access exists but MFA grant control not detected' -Severity 'High' -Evidence "Policies=$policyCount; MFA-grant-policy=False" -Recommendation 'Add/validate CA policy requiring MFA (at least for admins) and enforce strong methods.' -Effort 'Planned'
    }
}
catch {
    Write-Warn "Conditional Access policies not accessible: $($_)"
}

#==============================================================================
# C3. MFA Registration report
#==============================================================================
Write-Section "C3. MFA Registration (Best Effort)"
try {
    if ($IncludeMfaRegistrationReport) {
        $uri = "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
        if ($resp.value) {
            $mfa = $resp.value | Select-Object -Property @(
                @{Name='UserPrincipalName';Expression={$_.userPrincipalName}},
                @{Name='UserDisplayName';Expression={$_.userDisplayName}},
                @{Name='IsMfaCapable';Expression={$_.isMfaCapable}},
                @{Name='IsMfaRegistered';Expression={$_.isMfaRegistered}},
                @{Name='MethodsRegistered';Expression={($_.methodsRegistered -join '; ')}},
                @{Name='DefaultMfaMethod';Expression={$_.defaultMfaMethod}},
                @{Name='UserType';Expression={$_.userType}}
            )
            $mfa | Export-Csv -Path (Join-Path $outputRoot '22_MFA_Registration.csv') -NoTypeInformation -Encoding UTF8

            $member = $mfa | Where-Object { $_.UserType -eq 'Member' }
            $registeredPct = 0
            if ($member.Count -gt 0) {
                $registeredPct = [math]::Round((($member | Where-Object { $_.IsMfaRegistered -eq $true }).Count / $member.Count) * 100, 2)
            }
            $Scorecard['MFA Registered % (Members)'] = $registeredPct
            Write-Ok "MFA registered (members): $registeredPct%"

            if ($registeredPct -lt 90) {
                Add-Finding -Category 'Identity' -Finding 'MFA registration coverage below target' -Severity 'High' -Evidence "MFA registered (members) = $registeredPct%" -Recommendation 'Run MFA registration campaign and enforce via Conditional Access; prefer phishing-resistant methods for admins.' -Effort 'Planned'
            }
        }
    }
}
catch {
    Write-Warn "MFA registration report not available: $($_)"
}

#==============================================================================
# E. Shared Mailboxes (EXO) - Separate process
#==============================================================================
Write-Section "E. Shared Mailboxes (Exchange Online)"
if (-not $IncludeSharedMailboxes) {
    Write-Warn "Shared mailbox export skipped by parameter"
} else {
    $exoCsv = Join-Path $outputRoot '08_Shared_Mailboxes.csv'

    if ($RunEXOInSeparateProcess) {
        try {
            $exoScriptPath = Join-Path $outputRoot '_EXO_SharedMailbox_Export.ps1'
            $disableWamArg  = if($DisableWAMForEXO){ '-DisableWAM' } else { '' }

            $exoScript = @"
Import-Module ExchangeOnlineManagement
try {
    Connect-ExchangeOnline $disableWamArg -ShowBanner:`$false -ErrorAction Stop
    `$shared = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -PropertySets All |
        Select-Object DisplayName, PrimarySmtpAddress, Alias, UserPrincipalName, RecipientTypeDetails,
            @{Name='EmailAddresses';Expression={`$_.EmailAddresses -join '; '}},
            WhenCreated, WhenChanged, HiddenFromAddressListsEnabled,
            ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward,
            ArchiveStatus, RetentionPolicy, LitigationHoldEnabled,
            MaxSendSize, MaxReceiveSize

    `$shared | Export-Csv -Path '$exoCsv' -NoTypeInformation -Encoding UTF8
    Write-Host "[OK] Exported shared mailboxes: `$(`$shared.Count)" -ForegroundColor Green
}
catch {
    Write-Host "[ERR] EXO shared mailbox export failed: `$($_)" -ForegroundColor Red
}
finally {
    try { Disconnect-ExchangeOnline -Confirm:`$false -ErrorAction SilentlyContinue } catch {}
}
"@

            $exoScript | Set-Content -Path $exoScriptPath -Encoding UTF8

            $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
            if (-not $pwsh) { $pwsh = (Get-Command powershell -ErrorAction Stop).Source }

            Start-Process -FilePath $pwsh -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',"$exoScriptPath") -Wait | Out-Null

            if (Test-Path $exoCsv) {
                $count = (Import-Csv $exoCsv).Count
                $Scorecard['Shared Mailboxes'] = $count
                Write-Ok "Shared mailboxes exported: $count"
            } else {
                $Scorecard['Shared Mailboxes'] = 'Not exported'
                Write-Warn "Shared mailbox CSV not found - EXO login may have failed"
            }
        }
        catch {
            Write-Err "Shared mailbox export (separate process) failed: $($_)"
        }
    } else {
        Write-Warn "RunEXOInSeparateProcess is False. This may fail due to MSAL/WAM conflicts in some hosts."
    }
}

#==============================================================================
# F. Build Roadmap (30/60/90)
#==============================================================================
Write-Section "F. Build 30/60/90 Roadmap"
try {
    foreach ($f in ($Findings | Sort-Object @{Expression={Severity-Rank $_.Severity}}, Category)) {
        $bucket = '60 Days'
        if ($f.Severity -in @('Critical','High') -and $f.Effort -eq 'Quick') { $bucket = '30 Days' }
        elseif ($f.Severity -eq 'Medium') { $bucket = '90 Days' }
        elseif ($f.Severity -eq 'Low') { $bucket = '90 Days' }

        $Roadmap += [PSCustomObject]@{
            Timeline       = $bucket
            Category       = $f.Category
            WorkItem       = $f.Finding
            Recommendation = $f.Recommendation
            Evidence       = $f.Evidence
            Severity       = $f.Severity
        }
    }

    $Roadmap | Export-Csv -Path (Join-Path $outputRoot '40_Roadmap_30_60_90.csv') -NoTypeInformation -Encoding UTF8
    Write-Ok "Roadmap generated: $($Roadmap.Count) items"
}
catch {
    Write-Warn "Roadmap generation skipped: $($_)"
}

#==============================================================================
# G. Scorecard + Findings + Executive Summary
#==============================================================================
Write-Section "G. Export Scorecard + Findings + Executive Summary"
try {
    $scoreRows = @()
    foreach ($k in $Scorecard.Keys) {
        $scoreRows += [PSCustomObject]@{ Metric = $k; Value = $Scorecard[$k] }
    }
    $scoreRows | Export-Csv -Path (Join-Path $outputRoot '00_Scorecard.csv') -NoTypeInformation -Encoding UTF8

    $FindingsSorted = $Findings | Sort-Object @{Expression={Severity-Rank $_.Severity}}, Category
    $FindingsSorted | Export-Csv -Path (Join-Path $outputRoot '41_Findings_RiskRegister.csv') -NoTypeInformation -Encoding UTF8

    $topFindings = $FindingsSorted | Select-Object -First 10

    $execMd = @()
    $execMd += "# Microsoft 365 Assessment – Executive Summary"
    $execMd += ""
    $execMd += "**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $execMd += "**Report period (usage):** $ReportPeriod"
    $execMd += ""
    $execMd += "## Scorecard"
    foreach ($k in $Scorecard.Keys) { $execMd += "- **${k}:** $($Scorecard[$k])" }
    $execMd += ""
    $execMd += "## Top Priority Findings (Top 10)"

    if ($topFindings.Count -eq 0) {
        $execMd += "- No high-impact findings were generated by automated checks. Review Secure Score recommendations and validate with portal settings."
    } else {
        $i = 1
        foreach ($t in $topFindings) {
            $execMd += "$i. **[$($t.Severity)] $($t.Finding)** – $($t.Recommendation)"
            $i++
        }
    }

    $execMd += ""
    $execMd += "## 30/60/90-Day Remediation Roadmap"
    $execMd += "- **30 days:** Quick wins on critical controls (identity hardening, baseline protections)."
    $execMd += "- **60 days:** Implement high-impact policy controls (Conditional Access baselines, privileged access hygiene)."
    $execMd += "- **90 days:** Governance and optimization (license optimization, collaboration lifecycle controls, monitoring)."
    $execMd += ""
    $execMd += "## Notes"
    $execMd += "- Some data may not be accessible via Microsoft Graph in all tenants; validate in portals where required."
    $execMd += "- Severities and timelines are heuristics; align with customer constraints and change management." 

    $mdPath = Join-Path $outputRoot '00_Executive_Summary.md'
    $execMd | Out-File -FilePath $mdPath -Encoding UTF8

    $html = @()
    $html += "<html><head><meta charset='utf-8'><title>M365 Assessment Executive Summary</title>"
    $html += "<style>body{font-family:Segoe UI,Arial; margin:24px;} h1{color:#005a9e;} table{border-collapse:collapse;} td,th{border:1px solid #ddd;padding:8px;} th{background:#f3f6f9;} .sev-Critical{color:#b00020;font-weight:700;} .sev-High{color:#c75c00;font-weight:700;} .sev-Medium{color:#7a5c00;font-weight:700;} .sev-Low{color:#2b7a0b;font-weight:700;}</style>"
    $html += "</head><body>"
    $html += "<h1>Microsoft 365 Assessment – Executive Summary</h1>"
    $html += "<p><b>Generated:</b> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br><b>Report period (usage):</b> $ReportPeriod</p>"

    $html += "<h2>Scorecard</h2><table><tr><th>Metric</th><th>Value</th></tr>"
    foreach ($k in $Scorecard.Keys) { $html += "<tr><td>${k}</td><td>$($Scorecard[$k])</td></tr>" }
    $html += "</table>"

    $html += "<h2>Top Priority Findings (Top 10)</h2>"
    if ($topFindings.Count -eq 0) {
        $html += "<p>No high-impact findings were generated by automated checks. Review Secure Score recommendations and validate with portal settings.</p>"
    } else {
        $html += "<table><tr><th>#</th><th>Severity</th><th>Finding</th><th>Recommendation</th></tr>"
        $i = 1
        foreach ($t in $topFindings) {
            $sevClass = "sev-$($t.Severity)"
            $html += "<tr><td>$i</td><td class='$sevClass'>$($t.Severity)</td><td>$($t.Finding)</td><td>$($t.Recommendation)</td></tr>"
            $i++
        }
        $html += "</table>"
    }

    $html += "<h2>30/60/90-Day Remediation Roadmap</h2>"
    $html += "<p>See <b>40_Roadmap_30_60_90.csv</b> for the full action list.</p>"

    $html += "<h2>Notes</h2><ul>"
    $html += "<li>Some data may not be accessible via Microsoft Graph in all tenants; validate in portals where required.</li>"
    $html += "<li>Severities and timelines are heuristics; align with customer constraints and change management.</li>"
    $html += "</ul>"

    $html += "</body></html>"

    $htmlPath = Join-Path $outputRoot '00_Executive_Summary.html'
    $html | Out-File -FilePath $htmlPath -Encoding UTF8

    Write-Ok "Scorecard, Findings, Roadmap, and Executive Summary exported"
}
catch {
    Write-Err "Report generation failed: $($_)"
}

#==============================================================================
# DONE
#==============================================================================
Write-Section "DONE"
try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
Write-Ok "All outputs saved to: $outputRoot"
Invoke-Item $outputRoot
