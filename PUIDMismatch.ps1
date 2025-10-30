#################################################################
# CONFIG
#################################################################
$AdminSiteURL  = "https://oceancloudconsults-admin.sharepoint.com"
$SiteCollAdmin = "vipin@oceancloudconsults.onmicrosoft.com"
$AffectedUser  = "bhavprita@oceancloudconsults.onmicrosoft.com"
$ReportMode    = $true   # DRY RUN if $true

# === App-only (certificate) auth parameters ===
$Tenant        = "oceancloudconsults.onmicrosoft.com"
$ClientId      = "7147bbd7-b6d8-46c7-b6ae-d862c318c629"

# EITHER: use cert thumbprint in CurrentUser\My
$CertThumbprint = "ee349da6c000c7475af7caebfd9ab99843293393"
$UseThumbprint  = $true

# OR: use PFX file
# $PfxPath      = "C:\Secure\appcert.pfx"
# $PfxPassword  = ConvertTo-SecureString "P@ssw0rd!" -AsPlainText -Force
# $UseThumbprint = $false

# Processing batch size (client-side pagination while looping)
$BatchSize = 250

#################################################################
# REPORT & LOGS
#################################################################
Function Add-ReportRecord($SiteURL, $Action) {
    [pscustomobject]@{ "Site URL" = $SiteURL; "Action" = $Action } |
        Export-Csv -Path $ReportOutput -NoTypeInformation -Append
}
Function Add-ScriptLog($Color, $Msg) {
    Write-Host -ForegroundColor $Color $Msg
    $Date = Get-Date -Format "yyyy/MM/dd HH:mm"
    Add-Content -Path $LogsOutput -Value "$Date - $Msg"
}

# Create Report location
$FolderPath  = "$Env:USERPROFILE\Documents\"
$Date        = Get-Date -Format "yyyyMMddHHmmss"
$ReportName  = "IDMismatchSPO"
$FolderName  = "${Date}_$ReportName"
New-Item -Path $FolderPath -Name $FolderName -ItemType "directory" -Force | Out-Null
$ReportOutput = Join-Path (Join-Path $FolderPath $FolderName) ($FolderName + "_report.csv")
$LogsOutput   = Join-Path (Join-Path $FolderPath $FolderName) ($FolderName + "_Logs.txt")
Add-ScriptLog Cyan "Report will be generated at $ReportOutput"

#################################################################
# CONNECTION HELPERS
#################################################################
function Connect-Admin {
    if ($UseThumbprint) {
        Connect-PnPOnline -Url $AdminSiteURL -ClientId $ClientId -Tenant $Tenant `
            -Thumbprint $CertThumbprint -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $AdminSiteURL -ClientId $ClientId -Tenant $Tenant `
            -CertificatePath $PfxPath -CertificatePassword $PfxPassword -ErrorAction Stop
    }
}
function Connect-Site([string]$SiteUrl) {
    if ($UseThumbprint) {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant `
            -Thumbprint $CertThumbprint -ErrorAction Stop
    } else {
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Tenant $Tenant `
            -CertificatePath $PfxPath -CertificatePassword $PfxPassword -ErrorAction Stop
    }
}

#################################################################
# USER SNAPSHOT / REHYDRATE
#################################################################
function Get-UserAndMemberships {
    param([Parameter(Mandatory)][string]$UserEmail)
    $result = [ordered]@{ User=$null; Groups=@(); IsSiteAdmin=$false }
    $spUser = Get-PnPUser -ErrorAction SilentlyContinue | Where-Object { $_.Email -eq $UserEmail }
    if ($spUser) {
        $result.User        = $spUser
        $result.IsSiteAdmin = [bool]$spUser.IsSiteAdmin
        $groups = Get-PnPGroup -ErrorAction SilentlyContinue
        foreach ($g in $groups) {
            try {
                $members = Get-PnPGroupMembers -Identity $g -ErrorAction SilentlyContinue
                if ($members | Where-Object { $_.LoginName -eq $spUser.LoginName }) { $result.Groups += $g.Title }
            } catch { }
        }
    }
    return [PSCustomObject]$result
}
function Rehydrate-User {
    param(
        [Parameter(Mandatory)][string]$UserEmail,
        [Parameter(Mandatory)][string[]]$GroupsToRestore,
        [Parameter(Mandatory)][bool]$RestoreSCA,
        [Parameter(Mandatory)][string]$SiteUrl
    )
    $null = Ensure-PnPUser -LoginName $UserEmail -ErrorAction Stop
    foreach ($g in $GroupsToRestore) {
        try {
            Add-PnPGroupMember -Identity $g -Users $UserEmail -ErrorAction Stop
            Add-ScriptLog White "Restored group '$g' for $UserEmail"
            Add-ReportRecord -SiteURL $SiteUrl -Action "Restored group '$g' for $UserEmail"
        } catch {
            Add-ScriptLog DarkYellow "WARN: Failed to add $UserEmail to '$g' ($($_.Exception.Message))"
            Add-ReportRecord -SiteURL $SiteUrl -Action "WARN: Could not add to '$g' - $($_.Exception.Message)"
        }
    }
    if ($RestoreSCA) {
        try {
            Add-PnPSiteCollectionAdmin -Owners $UserEmail -ErrorAction Stop
            Add-ScriptLog White "Restored Site Collection Admin for $UserEmail"
            Add-ReportRecord -SiteURL $SiteUrl -Action "Restored Site Collection Admin for $UserEmail"
        } catch {
            Add-ScriptLog DarkYellow "WARN: Could not restore SCA for $UserEmail ($($_.Exception.Message))"
            Add-ReportRecord -SiteURL $SiteUrl -Action "WARN: Failed to restore SCA - $($_.Exception.Message)"
        }
    }
}

#################################################################
# APP CATALOG DETECTION (EXCLUDE)
#################################################################
# Will exclude tenant App Catalog site collection regardless of how the tenant reports it
$TenantAppCatalogUrl = $null
function Test-IsAppCatalogSite {
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter()][string]$Template
    )
    # 1) Exact match to tenant app catalog URL (if retrievable)
    if ($TenantAppCatalogUrl -and ($Url.TrimEnd('/') -ieq $TenantAppCatalogUrl.TrimEnd('/'))) { return $true }
    # 2) Template name sometimes reported as APPCATALOG#0
    if ($Template -and ($Template -like 'APPCATALOG*')) { return $true }
    # 3) Common URL patterns
    if ($Url -match '/(sites|teams)/appcatalog($|/)' -or $Url -match '/appcatalog($|/)$') { return $true }
    return $false
}

#################################################################
# MAIN
#################################################################

# Derive tenant root URLs to skip
# Example: https://oceancloudconsults-admin.sharepoint.com  -> tenant = oceancloudconsults
$tenantName = ([Uri]$AdminSiteURL).Host.Split('.')[0] -replace '-admin$',''
$SPO_Root   = "https://$tenantName.sharepoint.com/"
$ODB_Root   = "https://$tenantName-my.sharepoint.com/"

try {
    Connect-Admin
    Add-ScriptLog Cyan "Connected to SharePoint Admin Center"

    # Try get tenant App Catalog URL (may be null if not created)
    try { $TenantAppCatalogUrl = Get-PnPTenantAppCatalogUrl -ErrorAction Stop } catch { $TenantAppCatalogUrl = $null }

    # ---- Fetch STANDARD sites (PnP handles server-side paging internally) ----
    $stdSites = Get-PnPTenantSite | Where-Object {
        $_.Title -ne "" -and $_.Template -notlike "*Redirect*"
    }

    # ---- Fetch OneDrive PERSONAL sites (PnP handles server-side paging internally) ----
    $odbSites = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"

    # Combine & filter (skip SPO/ODB roots and App Catalog)
    $allSites = @($stdSites + $odbSites) |
        Where-Object {
            $_.Url -ne $SPO_Root -and
            $_.Url -ne $ODB_Root -and
            -not (Test-IsAppCatalogSite -Url $_.Url -Template $_.Template) -and
            $_.Url -and
            $_.Template -notlike "*Redirect*"
        } |
        Sort-Object Url -Unique

    Add-ScriptLog Cyan ("Collected sites - Standard: {0}, OneDrive: {1}, After filters: {2}" -f `
        $stdSites.Count, $odbSites.Count, $allSites.Count)
    Add-ReportRecord -SiteURL "-" -Action "TOTAL: Standard=$($stdSites.Count), ODB=$($odbSites.Count), Filtered=$($allSites.Count)"
}
catch {
    Add-ScriptLog Red "Admin connect/list error: $($_.Exception.Message)"
    break
}

# ---- Client-side batching for processing large tenants (5k+ sites safe) ----
$Total = $allSites.Count
$Processed = 0
for ($i = 0; $i -lt $Total; $i += $BatchSize) {

    $maxIndex = [Math]::Min($i + $BatchSize - 1, $Total - 1)
    $batch    = $allSites[$i..$maxIndex]

    Add-ScriptLog Yellow ("Processing batch {0}-{1} of {2} (BatchSize={3})" -f ($i+1), ($maxIndex+1), $Total, $BatchSize)

    foreach ($oSite in $batch) {
        $Processed++
        $PercentComplete = [math]::Round(($Processed / [math]::Max($Total,1)) * 100, 2)
        Add-ScriptLog Yellow "$PercentComplete% Completed - Processing Site Collection: $($oSite.Url)"

        try {
            # Replace owners so affected user isn't SCA (prevents 'cannot remove owners' warning)
            try {
                if (-not $ReportMode) {
                    Set-PnPTenantSite -Url $oSite.Url -Owners $SiteCollAdmin -ErrorAction Stop
                }
                Add-ReportRecord -SiteURL $oSite.Url -Action ($(if($ReportMode){"DRY-RUN: Would set owners to $SiteCollAdmin"}else{"Set owners to $SiteCollAdmin"}))
            } catch {
                Add-ScriptLog DarkYellow "WARN: Could not set owners on '$($oSite.Url)' ($($_.Exception.Message))"
                Add-ReportRecord -SiteURL $oSite.Url -Action "WARN: Set owners failed - $($_.Exception.Message)"
            }

            # Connect to site
            Connect-Site -SiteUrl $oSite.Url

            # OPTIONAL SAFETY: If this is the AffectedUser's own OneDrive, skip
            if ($oSite.Url -like "*-my.sharepoint.com/personal/*") {
                if ($oSite.Url -match "/personal/([^/]+)/") {
                    $ownerSegment = $Matches[1]
                    $affectedSegment = $AffectedUser.Split('@')[0].Replace('.','_')
                    if ($ownerSegment -like "*$affectedSegment*") {
                        Add-ScriptLog DarkGray "Detected personal ODB for the affected user. Skipping: $($oSite.Url)"
                        Add-ReportRecord -SiteURL $oSite.Url -Action "Skip: Affected user's own OneDrive"
                        continue
                    }
                }
            }

            # Snapshot current state
            $info = Get-UserAndMemberships -UserEmail $AffectedUser
            if (-not $info.User) {
                Add-ScriptLog DarkGray "User $AffectedUser not found in $($oSite.Url). Skipping."
                Add-ReportRecord -SiteURL $oSite.Url -Action "User not present. Skipped."
                continue
            }

            Add-ScriptLog White "Will repair identity for $AffectedUser in $($oSite.Url). Groups: $([string]::Join(', ',$info.Groups)) | SCA: $($info.IsSiteAdmin)"
            if ($ReportMode) {
                Add-ReportRecord -SiteURL $oSite.Url -Action "DRY-RUN: Would remove & re-add $AffectedUser; restore groups; SCA=$($info.IsSiteAdmin)"
                continue
            }

            # Try removing SCA first
            if ($info.IsSiteAdmin) {
                try {
                    Remove-PnPSiteCollectionAdmin -Owners $AffectedUser -ErrorAction Stop
                    Add-ScriptLog White "Removed SCA for $AffectedUser (pre-removal)"
                    Add-ReportRecord -SiteURL $oSite.Url -Action "Removed SCA (pre-removal) for $AffectedUser"
                } catch {
                    Add-ScriptLog DarkYellow "WARN: Could not remove SCA for $AffectedUser ($($_.Exception.Message))"
                    Add-ReportRecord -SiteURL $oSite.Url -Action "WARN: Remove SCA failed - $($_.Exception.Message)"
                }
            }

            # Remove user principal
            try {
                Remove-PnPUser -Identity $info.User.Id -Force -ErrorAction Stop
                Add-ScriptLog White "Removed user principal for $AffectedUser from $($oSite.Url)"
                Add-ReportRecord -SiteURL $oSite.Url -Action "Removed user principal ($AffectedUser)"
            } catch {
                Add-ScriptLog DarkYellow "WARN: Failed removing user $AffectedUser in $($oSite.Url) ($($_.Exception.Message))"
                Add-ReportRecord -SiteURL $oSite.Url -Action "WARN: Remove user failed - $($_.Exception.Message)"
                continue
            }

            # Rehydrate & restore
            Rehydrate-User -UserEmail $AffectedUser -GroupsToRestore $info.Groups -RestoreSCA $info.IsSiteAdmin -SiteUrl $oSite.Url
            Add-ScriptLog Green "Rehydrated identity and restored access for $AffectedUser in $($oSite.Url)"
            Add-ReportRecord -SiteURL $oSite.Url -Action "Rehydrated identity and restored access for $AffectedUser"
        }
        catch {
            Add-ScriptLog Red "Error while processing '$($oSite.Url)'"
            Add-ScriptLog Red "Error message: '$($_.Exception.Message)'"
            Add-ScriptLog Red "Error trace: '$($_.Exception.ScriptStackTrace)'"
            Add-ReportRecord -SiteURL $oSite.Url -Action $_.Exception.Message
        }
    }
}

Add-ScriptLog Cyan "100% Completed - Finished running script"
Add-ScriptLog Cyan "Report generated at $ReportOutput"
