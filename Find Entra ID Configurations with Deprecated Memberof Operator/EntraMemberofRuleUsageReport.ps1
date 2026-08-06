<#
=============================================================================================

Name         : Find Entra ID Configurations That Still Use the Deprecated memberOf Operator 
Version      : 1.0
Website      : blog.admindroid.com



For detailed script execution: 

============================================================================================
#>

[CmdletBinding()]
Param(
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [Switch]$DynamicGroups,
    [Switch]$AdministrativeUnits,
    [Switch]$AutoAssignmentPolicies,
    [Switch]$Unattended
)

# Connect to Microsoft Graph

Function Connect-MgGraphSession {
    $MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
    if ($null -eq $MsGraphModule) {
        Write-Host "`nImportant: Microsoft Graph module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $Confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph module..."
            Install-Module Microsoft.Graph -Scope CurrentUser -AllowClobber
            Write-Host "Microsoft Graph module is installed in the machine successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Exiting.`nNote: Microsoft Graph module must be available in your system to run the script." -ForegroundColor Red
            Exit
        }
    }

    #Disconnect existing session
    if ($null -ne (Get-MgContext)) {
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "`nConnecting to Microsoft Graph..."

    if (($TenantId -ne "") -and ($AppId -ne "") -and ($CertificateThumbprint -ne "")) {
        Connect-MgGraph -TenantId $TenantId -AppId $AppId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        if ($null -ne (Get-MgContext)) {
            Write-Host "Connected to Microsoft Graph PowerShell using $(((Get-MgContext).AppName)) application certificate." -ForegroundColor Green
        }
    }
    else {
        Connect-MgGraph -Scopes "Group.Read.All", "AdministrativeUnit.Read.All", "EntitlementManagement.Read.All" -NoWelcome
        if ($null -ne (Get-MgContext)) {
            Write-Host "Connected to Microsoft Graph PowerShell using $((Get-MgContext).Account) account.`n" -ForegroundColor Green
        }
    }

    if ($null -eq (Get-MgContext)) {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
}

# Helpers

Function Get-FileTimestamp {
    return (Get-Date -Format 'yyyy-MMM-dd-ddd HH-mm-ss').ToString()
}

Function Get-MembershipRuleFromTarget {
    Param($Target)

    # Typed SDK objects carry the rule in AdditionalProperties, raw hashtables expose it directly.
    if ($Target.AdditionalProperties -and $Target.AdditionalProperties.ContainsKey('membershipRule')) {
        return $Target.AdditionalProperties['membershipRule']
    }
    return $Target.membershipRule
}

# Dynamic group processing

Function Process-DynamicGroups {
    $Group = $_
    $script:GroupProcessedCount++
    $DisplayName = $Group.DisplayName
    Write-Progress -Activity "Scanning dynamic membership groups" `
        -Status "Processing $DisplayName (group $script:GroupProcessedCount)"

    $MembershipRule = $Group.MembershipRule
    if ($MembershipRule -notmatch 'memberOf') { return }

    if ($Group.GroupTypes -contains 'Unified') {
        $GroupType = "Microsoft 365"
    }
    elseif ($Group.SecurityEnabled) {
        $GroupType = "Security"
    }
    else {
        $GroupType = "Distribution"
    }

    $Values = @(
        ($DisplayName -replace '"', '""'),
        $Group.Id,
        $GroupType,
        $Group.MembershipRuleProcessingState,
        ($MembershipRule -replace '"', '""')
    )

    if ($null -eq $script:GroupWriter) {
        $script:GroupWriter = New-Object System.IO.StreamWriter -ArgumentList $script:GroupCSV, $false, ([System.Text.Encoding]::UTF8)
        $script:GroupWriter.AutoFlush = $true
        $script:GroupWriter.WriteLine($script:GroupCsvHeader)
    }
    $script:GroupWriter.WriteLine('"' + ($Values -join '","') + '"')
    $script:GroupCount++
}

# Administrative unit processing

Function Process-AdministrativeUnits {
    $AdminUnit = $_
    $script:AdminUnitProcessedCount++
    $DisplayName = $AdminUnit.DisplayName
    Write-Progress -Activity "Scanning dynamic administrative units" `
        -Status "Processing $DisplayName (administrative unit $script:AdminUnitProcessedCount)"

    if ($AdminUnit.MembershipType -ne 'Dynamic') { return }

    $MembershipRule = $AdminUnit.MembershipRule
    if ($MembershipRule -notmatch 'memberOf') { return }

    $Values = @(
        ($DisplayName -replace '"', '""'),
        $AdminUnit.Id,
        $AdminUnit.MembershipRuleProcessingState,
        ($MembershipRule -replace '"', '""')
    )

    if ($null -eq $script:AdminUnitWriter) {
        $script:AdminUnitWriter = New-Object System.IO.StreamWriter -ArgumentList $script:AdminUnitCSV, $false, ([System.Text.Encoding]::UTF8)
        $script:AdminUnitWriter.AutoFlush = $true
        $script:AdminUnitWriter.WriteLine($script:AdminUnitCsvHeader)
    }
    $script:AdminUnitWriter.WriteLine('"' + ($Values -join '","') + '"')
    $script:AdminUnitCount++
}

# Auto-assignment policy processing

Function Process-AutoAssignmentPolicies {
    $Policy = $_
    $script:PolicyProcessedCount++
    $DisplayName = $Policy.DisplayName
    Write-Progress -Activity "Scanning entitlement management auto-assignment policies" `
        -Status "Processing $DisplayName (policy $script:PolicyProcessedCount)"

    # Only auto-assignment policies populate specificAllowedTargets.
    $Targets = $Policy.SpecificAllowedTargets
    if ($null -eq $Targets) {
        $Targets = $Policy.AdditionalProperties['specificAllowedTargets']
    }
    if ($null -eq $Targets) { return }

    foreach ($Target in $Targets) {
        $MembershipRule = Get-MembershipRuleFromTarget -Target $Target
        if ($MembershipRule -notmatch 'memberOf') { continue }

        $Values = @(
            ($DisplayName -replace '"', '""'),
            $Policy.Id,
            (([string]$Policy.AccessPackage.DisplayName) -replace '"', '""'),
            $Policy.AccessPackage.Id,
            ($MembershipRule -replace '"', '""')
        )

        if ($null -eq $script:PolicyWriter) {
            $script:PolicyWriter = New-Object System.IO.StreamWriter -ArgumentList $script:PolicyCSV, $false, ([System.Text.Encoding]::UTF8)
            $script:PolicyWriter.AutoFlush = $true
            $script:PolicyWriter.WriteLine($script:PolicyCsvHeader)
        }
        $script:PolicyWriter.WriteLine('"' + ($Values -join '","') + '"')
        $script:PolicyCount++
    }
}

# Report setup

$Location = Get-Location
$Timestamp = Get-FileTimestamp
$script:GroupCSV = "$Location\MemberOf_DynamicGroups_$Timestamp.csv"
$script:AdminUnitCSV = "$Location\MemberOf_AdministrativeUnits_$Timestamp.csv"
$script:PolicyCSV = "$Location\MemberOf_AutoAssignmentPolicies_$Timestamp.csv"

$script:GroupCsvHeader = '"Group Name","Group Id","Group Type","Membership Rule Processing State","Membership Rule"'
$script:AdminUnitCsvHeader = '"Administrative Unit Name","Administrative Unit Id","Membership Rule Processing State","Membership Rule"'
$script:PolicyCsvHeader = '"Policy Name","Policy Id","Access Package Name","Access Package Id","Membership Rule"'

$script:GroupWriter = $null
$script:AdminUnitWriter = $null
$script:PolicyWriter = $null
$script:GroupCount = 0
$script:AdminUnitCount = 0
$script:PolicyCount = 0
$script:GroupProcessedCount = 0
$script:AdminUnitProcessedCount = 0
$script:PolicyProcessedCount = 0

# Surface selection

$ScanGroups = $DynamicGroups.IsPresent
$ScanAdminUnits = $AdministrativeUnits.IsPresent
$ScanPolicies = $AutoAssignmentPolicies.IsPresent
if (-not ($ScanGroups -or $ScanAdminUnits -or $ScanPolicies)) {
    $ScanGroups = $true
    $ScanAdminUnits = $true
    $ScanPolicies = $true
}

Connect-MgGraphSession

# Stream rows.
try {
    if ($ScanGroups) {
        Write-Host "Checking dynamic membership groups for memberOf rules..." -ForegroundColor Cyan
        Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'DynamicMembership')" `
            -Property "id,displayName,groupTypes,securityEnabled,membershipRule,membershipRuleProcessingState" |
            ForEach-Object { Process-DynamicGroups }
        Write-Progress -Activity "Scanning dynamic membership groups" -Completed
    }

    if ($ScanAdminUnits) {
        Write-Host "Checking dynamic administrative units for memberOf rules..." -ForegroundColor Cyan
        Get-MgDirectoryAdministrativeUnit -All `
            -Property "id,displayName,membershipType,membershipRule,membershipRuleProcessingState" |
            ForEach-Object { Process-AdministrativeUnits }
        Write-Progress -Activity "Scanning dynamic administrative units" -Completed
    }

    if ($ScanPolicies) {
        Write-Host "Checking entitlement management auto-assignment policies for memberOf rules..." -ForegroundColor Cyan
        Get-MgEntitlementManagementAssignmentPolicy -All -ExpandProperty "accessPackage" |
            ForEach-Object { Process-AutoAssignmentPolicies }
        Write-Progress -Activity "Scanning entitlement management auto-assignment policies" -Completed
    }
}
finally {
    if ($null -ne $script:GroupWriter) {
        $script:GroupWriter.Close()
        $script:GroupWriter.Dispose()
    }
    if ($null -ne $script:AdminUnitWriter) {
        $script:AdminUnitWriter.Close()
        $script:AdminUnitWriter.Dispose()
    }
    if ($null -ne $script:PolicyWriter) {
        $script:PolicyWriter.Close()
        $script:PolicyWriter.Dispose()
    }
}

# Output

$GeneratedFiles = @()

Write-Host "`nConfigurations using the memberOf operator:" -ForegroundColor Cyan
if ($ScanGroups) {
    Write-Host "Dynamic membership groups                : $script:GroupCount"
    if ($script:GroupCount -gt 0) { $GeneratedFiles += $script:GroupCSV }
}
if ($ScanAdminUnits) {
    Write-Host "Dynamic administrative units             : $script:AdminUnitCount"
    if ($script:AdminUnitCount -gt 0) { $GeneratedFiles += $script:AdminUnitCSV }
}
if ($ScanPolicies) {
    Write-Host "Entitlement management policies          : $script:PolicyCount"
    if ($script:PolicyCount -gt 0) { $GeneratedFiles += $script:PolicyCSV }
}

$TotalCount = $script:GroupCount + $script:AdminUnitCount + $script:PolicyCount
Write-Host "`nSummary: $TotalCount configurations must be migrated before November 3, 2026." -ForegroundColor Yellow

if ($GeneratedFiles.Count -gt 0) {
    Write-Host "`nThe reports are available in:" -ForegroundColor Yellow
    foreach ($GeneratedFile in $GeneratedFiles) {
        Write-Host $GeneratedFile
    }
}


Write-Host "`n~~ The Script is prepared by AdminDroid Community ~~`n" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
Write-Host "admindroid.com" -NoNewline -ForegroundColor Yellow
Write-Host " to get access to 3500+ Microsoft 365 reports & 450+ management actions. ~~`n" -ForegroundColor Green

if ($GeneratedFiles.Count -gt 0 -and -not $Unattended) {
    $Prompt = New-Object -ComObject wscript.shell
    $UserInput = $Prompt.popup("Do you want to open the output files?", 0, "Open Output Files", 4)
    if ($UserInput -eq 6) {
        foreach ($GeneratedFile in $GeneratedFiles) {
            Invoke-Item $GeneratedFile
        }
    }
}

Disconnect-MgGraph | Out-Null
