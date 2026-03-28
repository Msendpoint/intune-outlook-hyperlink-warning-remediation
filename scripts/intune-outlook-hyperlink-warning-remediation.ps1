<#
.SYNOPSIS
    Detects and remediates Outlook hyperlink security warnings and configures trusted sites via registry settings.

.DESCRIPTION
    This script is designed to be deployed as an Intune Proactive Remediation package (Detection + Remediation).
    It targets two common pain points in enterprise environments:
      1. Microsoft Outlook and Office applications displaying intrusive security warnings when users
         click on hyperlinks pointing to trusted internal resources (file servers, SharePoint, etc.).
      2. Missing or incorrect ZoneMap registry entries that define trusted sites for Internet Explorer
         and legacy Office security zones.

    The DETECTION script checks:
      - Whether each trusted site is registered in the ZoneMap registry key under HKLM Policies.
      - Whether the 'DisableHyperlinkWarning' DWORD is set to 1 in both the Outlook and Office
        security registry paths.
    Exits 0 (Compliant) or 1 (Non-Compliant) so Intune can trigger remediation accordingly.

    The REMEDIATION script:
      - Creates the ZoneMap registry path if missing and adds/updates all trusted site entries.
      - Creates the Outlook and Office security registry paths if missing.
      - Sets 'DisableHyperlinkWarning' to 1 in both paths to suppress security prompts.

    HOW TO DEPLOY:
      1. Save the Detection block and Remediation block as separate .ps1 files.
      2. In the Intune Admin Center, navigate to:
         Devices > Manage devices > Scripts and remediations > Create script package.
      3. Upload the Detection script and Remediation script, assign to target device groups.
      4. Monitor results under Devices > Manage devices > Scripts and remediations.

    CUSTOMIZATION:
      Update the $trustedSites hashtable in both scripts to reflect your organization's internal
      hostnames, file server UNC paths, and SharePoint URLs before deployment.

.NOTES
    Author:      Souhaiel Morhag
    Company:     MSEndpoint.com
    Blog:        https://msendpoint.com
    Academy:     https://app.msendpoint.com/academy
    LinkedIn:    https://linkedin.com/in/souhaiel-morhag
    GitHub:      https://github.com/Msendpoint
    License:     MIT

.EXAMPLE
    # Run the detection script manually to check compliance on a local machine:
    .\Detect-OutlookHyperlinkWarning.ps1

    # Run the remediation script manually to apply registry changes:
    .\Remediate-OutlookHyperlinkWarning.ps1
#>

[CmdletBinding()]
param (
    # Set to 'Detect' to run detection logic only, or 'Remediate' to apply fixes.
    [ValidateSet('Detect', 'Remediate')]
    [string]$Mode = 'Detect'
)

#region --- Configuration ---

# Registry path for ZoneMap trusted sites (Machine-level, enforced via policy)
$zoneMapPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains'

# Registry path for Outlook 2016/365 security settings (per-user)
$outlookSecurityPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Security'

# Registry path for Office-wide security settings (per-user)
$officeSecurityPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\Security'

# Trusted sites hashtable: Key = hostname or UNC path, Value = Zone (1 = Intranet, 2 = Trusted)
# !! CUSTOMIZE these entries to match your environment before deploying !!
$trustedSites = @{
    'file://fileservername.domain' = 1
    'file://internal.local'        = 1
    'server.internal.local'        = 1
    'company.sharepoint.com'       = 1
}

#endregion

#region --- Detection Logic ---

function Invoke-Detection {
    [CmdletBinding()]
    param()

    $compliance = $true

    Write-Verbose 'Starting compliance detection...'

    # --- Check trusted sites in ZoneMap ---
    foreach ($site in $trustedSites.Keys) {
        try {
            $existingValue = Get-ItemProperty -Path $zoneMapPath -Name $site -ErrorAction Stop
            if ($existingValue.$site -ne $trustedSites[$site]) {
                Write-Verbose "ZoneMap entry mismatch for '$site'. Expected: $($trustedSites[$site]), Found: $($existingValue.$site)"
                $compliance = $false
            }
            else {
                Write-Verbose "ZoneMap entry OK for '$site'."
            }
        }
        catch {
            Write-Verbose "ZoneMap entry missing for '$site'."
            $compliance = $false
        }
    }

    # --- Check Outlook DisableHyperlinkWarning ---
    try {
        $outlookHyperlinkWarning = Get-ItemProperty -Path $outlookSecurityPath -Name 'DisableHyperlinkWarning' -ErrorAction Stop
        if ($outlookHyperlinkWarning.DisableHyperlinkWarning -ne 1) {
            Write-Verbose "Outlook DisableHyperlinkWarning is NOT set to 1. Current value: $($outlookHyperlinkWarning.DisableHyperlinkWarning)"
            $compliance = $false
        }
        else {
            Write-Verbose 'Outlook DisableHyperlinkWarning is correctly set to 1.'
        }
    }
    catch {
        Write-Verbose 'Outlook DisableHyperlinkWarning registry value is missing.'
        $compliance = $false
    }

    # --- Check Office-wide DisableHyperlinkWarning ---
    try {
        $officeHyperlinkWarning = Get-ItemProperty -Path $officeSecurityPath -Name 'DisableHyperlinkWarning' -ErrorAction Stop
        if ($officeHyperlinkWarning.DisableHyperlinkWarning -ne 1) {
            Write-Verbose "Office DisableHyperlinkWarning is NOT set to 1. Current value: $($officeHyperlinkWarning.DisableHyperlinkWarning)"
            $compliance = $false
        }
        else {
            Write-Verbose 'Office DisableHyperlinkWarning is correctly set to 1.'
        }
    }
    catch {
        Write-Verbose 'Office DisableHyperlinkWarning registry value is missing.'
        $compliance = $false
    }

    # --- Report result ---
    if ($compliance) {
        Write-Output 'Compliant'
        exit 0
    }
    else {
        Write-Output 'Non-Compliant'
        exit 1
    }
}

#endregion

#region --- Remediation Logic ---

function Invoke-Remediation {
    [CmdletBinding()]
    param()

    Write-Verbose 'Starting remediation...'

    # --- Ensure ZoneMap registry path exists ---
    try {
        if (-not (Test-Path -Path $zoneMapPath)) {
            Write-Verbose "Creating ZoneMap path: $zoneMapPath"
            New-Item -Path $zoneMapPath -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Error "Failed to create ZoneMap registry path '$zoneMapPath'. Error: $_"
        exit 1
    }

    # --- Add or update trusted sites in ZoneMap ---
    foreach ($site in $trustedSites.Keys) {
        try {
            New-ItemProperty -Path $zoneMapPath -Name $site -Value $trustedSites[$site] `
                -PropertyType DWord -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Set ZoneMap entry: '$site' = $($trustedSites[$site])"
        }
        catch {
            Write-Warning "Failed to set ZoneMap entry for '$site'. Error: $_"
        }
    }

    # --- Ensure Outlook security path exists and set DisableHyperlinkWarning ---
    try {
        if (-not (Test-Path -Path $outlookSecurityPath)) {
            Write-Verbose "Creating Outlook security registry path: $outlookSecurityPath"
            New-Item -Path $outlookSecurityPath -Force -ErrorAction Stop | Out-Null
        }
        New-ItemProperty -Path $outlookSecurityPath -Name 'DisableHyperlinkWarning' -Value 1 `
            -PropertyType DWord -Force -ErrorAction Stop | Out-Null
        Write-Verbose 'Outlook DisableHyperlinkWarning set to 1.'
    }
    catch {
        Write-Error "Failed to configure Outlook hyperlink warning setting. Error: $_"
        exit 1
    }

    # --- Ensure Office security path exists and set DisableHyperlinkWarning ---
    try {
        if (-not (Test-Path -Path $officeSecurityPath)) {
            Write-Verbose "Creating Office security registry path: $officeSecurityPath"
            New-Item -Path $officeSecurityPath -Force -ErrorAction Stop | Out-Null
        }
        New-ItemProperty -Path $officeSecurityPath -Name 'DisableHyperlinkWarning' -Value 1 `
            -PropertyType DWord -Force -ErrorAction Stop | Out-Null
        Write-Verbose 'Office DisableHyperlinkWarning set to 1.'
    }
    catch {
        Write-Error "Failed to configure Office hyperlink warning setting. Error: $_"
        exit 1
    }

    Write-Output 'Remediation applied successfully.'
    exit 0
}

#endregion

#region --- Entry Point ---

switch ($Mode) {
    'Detect'    { Invoke-Detection }
    'Remediate' { Invoke-Remediation }
}

#endregion
