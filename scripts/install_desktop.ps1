<#
.SYNOPSIS
Install the self-contained RandomOffice executable for the current Windows user.
.DESCRIPTION
Creates a stable local install, desktop/Start menu shortcuts and Windows Open
With / Default apps registrations. Existing default choices remain unchanged.
Use -WhatIf to review all filesystem and registry operations without writing.
.EXAMPLE
.\scripts\install_desktop.ps1 -WhatIf
.EXAMPLE
.\scripts\install_desktop.ps1 -OpenDefaultApps
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [string]$SourceExe = '',
    [Parameter()]
    [switch]$OpenDefaultApps
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($SourceExe)) {
    # Also works when this script is distributed beside the standalone EXE.
    $SourceExe = Join-Path $PSScriptRoot 'RandomOffice.exe'
    if (-not (Test-Path -LiteralPath $SourceExe -PathType Leaf)) {
        $SourceExe = Join-Path $PSScriptRoot '..\dist\standalone\RandomOffice.exe'
    }
}
if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
    throw 'RandomOffice desktop installation requires Windows.'
}
if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    throw 'LOCALAPPDATA is not available for the current user.'
}

$source = Get-Item -LiteralPath $SourceExe
if ($source.PSIsContainer -or $source.Extension -ine '.exe') {
    throw 'SourceExe must point to the self-contained RandomOffice.exe file.'
}
if (Test-Path -LiteralPath (Join-Path $source.DirectoryName '_internal')) {
    throw 'This looks like an older multi-file build. Use dist\standalone\RandomOffice.exe instead.'
}
$fileStream = [System.IO.File]::OpenRead($source.FullName)
try {
    if ($fileStream.ReadByte() -ne 0x4D -or $fileStream.ReadByte() -ne 0x5A) {
        throw 'The source is not a Windows executable (MZ header missing).'
    }
} finally {
    $fileStream.Dispose()
}

$installDirectory = [System.IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'RandomOffice'))
$installedExe = Join-Path $installDirectory 'RandomOffice.exe'
$desktopDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::DesktopDirectory)
$programsDirectory = [Environment]::GetFolderPath([Environment+SpecialFolder]::Programs)
if ([string]::IsNullOrWhiteSpace($desktopDirectory) -or [string]::IsNullOrWhiteSpace($programsDirectory)) {
    throw 'Windows Desktop or Start menu directory could not be resolved.'
}

Write-Host "Install directory: $installDirectory"
Write-Host "Source: $($source.FullName)"
Write-Host 'Existing Word, Excel and PowerPoint default app choices are preserved.'

if ($source.FullName -ine $installedExe) {
    if ($PSCmdlet.ShouldProcess($installedExe, 'Install or update the self-contained application')) {
        [System.IO.Directory]::CreateDirectory($installDirectory) | Out-Null
        # Stage and verify before replacing the stable executable. A running app
        # will cause replacement to fail before any registrations are changed.
        $stagedExe = Join-Path $installDirectory ('RandomOffice-' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            Copy-Item -LiteralPath $source.FullName -Destination $stagedExe
            if ((Get-FileHash -LiteralPath $stagedExe -Algorithm SHA256).Hash -ne
                (Get-FileHash -LiteralPath $source.FullName -Algorithm SHA256).Hash) {
                throw 'The copied executable did not pass SHA256 verification.'
            }
            Move-Item -LiteralPath $stagedExe -Destination $installedExe -Force
            Unblock-File -LiteralPath $installedExe -ErrorAction SilentlyContinue
        } finally {
            if (Test-Path -LiteralPath $stagedExe) {
                Remove-Item -LiteralPath $stagedExe -Force
            }
        }
    }
}

$command = '"' + $installedExe + '" "%1"'
$icon = '"' + $installedExe + '",0'
$capabilities = 'Software\RandomOffice\Capabilities'
$application = 'Software\Classes\Applications\RandomOffice.exe'
$fileTypes = @(
    @{ Extension = '.docx'; ProgId = 'RandomOffice.Document'; Description = 'RandomOffice Word Document' },
    @{ Extension = '.xlsx'; ProgId = 'RandomOffice.Workbook'; Description = 'RandomOffice Excel Workbook' },
    @{ Extension = '.pptx'; ProgId = 'RandomOffice.Presentation'; Description = 'RandomOffice PowerPoint Presentation' }
)

# Each entry is a string value under HKCU. Empty Name means (Default).
# .ext keys receive only OpenWithProgids; never change their default value.
$registrations = @(
    @{ Key = $capabilities; Name = 'ApplicationName'; Value = 'RandomOffice' },
    @{ Key = $capabilities; Name = 'ApplicationDescription'; Value = 'Create, edit and save Word documents, Excel workbooks and PowerPoint presentations locally.' },
    @{ Key = $capabilities; Name = 'ApplicationIcon'; Value = $icon },
    @{ Key = 'Software\RegisteredApplications'; Name = 'RandomOffice'; Value = $capabilities },
    @{ Key = $application; Name = 'FriendlyAppName'; Value = 'RandomOffice' },
    @{ Key = "$application\DefaultIcon"; Name = ''; Value = $icon },
    @{ Key = "$application\shell\open\command"; Name = ''; Value = $command },
    @{ Key = 'Software\Microsoft\Windows\CurrentVersion\App Paths\RandomOffice.exe'; Name = ''; Value = $installedExe }
)
foreach ($fileType in $fileTypes) {
    $progidKey = 'Software\Classes\' + $fileType.ProgId
    $registrations += @(
        @{ Key = $progidKey; Name = ''; Value = $fileType.Description },
        @{ Key = $progidKey; Name = 'FriendlyTypeName'; Value = $fileType.Description },
        @{ Key = "$progidKey\DefaultIcon"; Name = ''; Value = $icon },
        @{ Key = "$progidKey\shell\open\command"; Name = ''; Value = $command },
        @{ Key = ('Software\Classes\' + $fileType.Extension + '\OpenWithProgids'); Name = $fileType.ProgId; Value = '' },
        @{ Key = "$application\SupportedTypes"; Name = $fileType.Extension; Value = '' },
        @{ Key = "$capabilities\FileAssociations"; Name = $fileType.Extension; Value = $fileType.ProgId }
    )
}

foreach ($entry in $registrations) {
    $displayName = if ($entry.Name) { $entry.Name } else { '(Default)' }
    $target = 'HKCU\' + $entry.Key + ' [' + $displayName + '] = ' + $entry.Value
    if ($PSCmdlet.ShouldProcess($target, 'Register RandomOffice as an available app')) {
        $registryKey = [Microsoft.Win32.Registry]::CurrentUser.CreateSubKey($entry.Key)
        try {
            $registryKey.SetValue($entry.Name, $entry.Value, [Microsoft.Win32.RegistryValueKind]::String)
        } finally {
            $registryKey.Dispose()
        }
    }
}

$shortcuts = @(
    (Join-Path $desktopDirectory 'RandomOffice.lnk'),
    (Join-Path $programsDirectory 'RandomOffice.lnk')
)
foreach ($shortcutPath in $shortcuts) {
    if ($PSCmdlet.ShouldProcess($shortcutPath, "Create shortcut to $installedExe")) {
        [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($shortcutPath)) | Out-Null
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $null
        try {
            $shortcut = $shell.CreateShortcut($shortcutPath)
            # COM shortcut properties contain raw paths, without command quotes.
            $shortcut.TargetPath = $installedExe
            $shortcut.WorkingDirectory = $installDirectory
            $shortcut.IconLocation = $icon
            $shortcut.Description = 'RandomOffice - Word, Excel and PowerPoint'
            $shortcut.WindowStyle = 3
            $shortcut.Save()
        } finally {
            if ($null -ne $shortcut) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut) }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
        }
    }
}

if ($PSCmdlet.ShouldProcess('Windows Shell association cache', 'Refresh available applications and document icons')) {
    if (-not ('RandomOffice.InstallerShell' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace RandomOffice {
    public static class InstallerShell {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(int eventId, uint flags, IntPtr item1, IntPtr item2);
    }
}
'@
    }
    # SHCNE_ASSOCCHANGED and SHCNF_IDLIST, with unused item pointers NULL.
    [RandomOffice.InstallerShell]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

if (-not $WhatIfPreference) {
    Write-Host 'RandomOffice installed. Open it from the desktop or Start menu.'
    Write-Host 'To choose it for Office documents: Settings > Apps > Default apps > RandomOffice.'
}
if ($OpenDefaultApps) {
    $settingsUri = 'ms-settings:defaultapps'
    if ([Environment]::OSVersion.Version.Build -ge 22000) {
        $settingsUri += '?registeredAppUser=RandomOffice'
    }
    if ($PSCmdlet.ShouldProcess($settingsUri, 'Open Windows Default apps settings')) {
        Start-Process -FilePath $settingsUri -WindowStyle Hidden
    }
}
