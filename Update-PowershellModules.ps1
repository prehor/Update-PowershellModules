# code: insertSpaces=true tabSize=2

<#
Cypyright (c) Petr Řehoř https://github.com/prehor. All rights reserved.
Licensed under the Apache License, Version 2.0.
#>

<#
.SYNOPSIS
Update installed PowerShell modules.

.DESCRIPTION
Update installed PowerShell modules from PowerShell Gallery.

.PARAMETER Name
Module name. Accept wildcard characters.

.LINK
https://github.com/prehor/Update-PowershellModules
https://www.powershellgallery.com
#>

###############################################################################
### PARAMETERS ################################################################
###############################################################################

#region Parameters

[CmdletBinding()]
param(
  [Parameter()]
  [String[]]$Name = @('*')
)

Set-StrictMode -Version Latest

#endregion

###############################################################################
### MAIN ######################################################################
###############################################################################

#region Main

# Upgrade installed modules
Get-InstalledModule -Name $Name |
Sort-Object -Property Name |
ForEach-Object -Begin {
  # Obsoleted versions with dependencies to be uninstalled later
  [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
  $ObsoletedVersions = @()
} -Process {
  $InstalledModule = $_
  # Get all installed versions
  $InstalledVersions = @(
    Get-InstalledModule -Name $InstalledModule.Name -AllVersions |
    Sort-Object -Property Vesion
  )
  # Get latest version
  try {
    $LatestModule = Find-Module -Name $InstalledModule.Name -ErrorAction SilentlyContinue
  } catch [Microsoft.PowerShell.PackageManagement.Cmdlets.FindPackage.NoMatchFoundForCriteria] {
    Write-Host -ForegroundColor Red "$($InstalledModule.Name) - The latest version $($InstalledModule.Version) ($($InstalledModule.PublishedDate)) not found in modules repository!"
    $LatestModule = $InstalledModule
  }
  # Upgrade module if needed
  if ($InstalledVersions.Version -NotContains $LatestModule.Version) {
    Write-Host -ForegroundColor Yellow "$($InstalledModule.Name) - Upgrading to $($LatestModule.Version) ($($LatestModule.PublishedDate))"
    Update-Module -Name $InstalledModule.Name -Force
  } else {
    Write-Host -ForegroundColor Green "$($InstalledModule.Name) - The latest version $($LatestModule.Version) ($($LatestModule.PublishedDate)) is already installed"
  }
  # Uninstall obsoleted versions
  try {
    $InstalledVersions |
    Where-Object { $_.Version -ne $LatestModule.Version } |
    Sort-Object -Property Version |
    ForEach-Object {
      $ObsoletedVersion = $_
      Write-Host -ForegroundColor Cyan "$($ObsoletedVersion.Name) - Uninstalling obsoleted version $($ObsoletedVersion.Version) ($($ObsoletedVersion.PublishedDate))"
      $ObsoletedVersion | Uninstall-Module -Force
    }
  } catch {
    # Catch obsoleted version with dependencies to be uninstalled later
    $ObsoletedVersions += $ObsoletedVersion
  }
} -End {
  # Recursively uninstall obsoleted versions with dependencies
  while ($ObsoletedVersions) {
    $RemainingModuleVersions = @()
    $ObsoletedVersions |
    ForEach-Object {
      $ObsoletedVersion = $_
      try {
        Write-Host -ForegroundColor Cyan "$($ObsoletedVersion.Name) - Uninstalling obsoleted version $($ObsoletedVersion.Version) ($($ObsoletedVersion.PublishedDate))"
        $ObsoletedVersion | Uninstall-Module -Force -ErrorAction Continue
      } catch {
        # Catch module versions which still have dependencies to be uninstalled later
        $RemainingModuleVersions += $ObsoletedVersion
      }
    }
    # No module versions were uninstalled due to dependencies on modules outside the updated module name scope
    if ($RemainingModuleVersions.Count -eq $ObsoletedVersions.Count) {
      $RemainingModuleVersions |
      ForEach-Object {
        $ObsoletedVersion = $_
        Write-Host -ForegroundColor Red "$($ObsoletedVersion.Name) - Cannot uninstall obsoleted version $($ObsoletedVersion.Version) ($($ObsoletedVersion.PublishedDate)) because of dependencies!"
      }
      # Stop uninstalling modules with dependencies on modules outside the updated module name scope
      $RemainingModuleVersions = @()
    }
    # Uninstall the remaining modules
    $ObsoletedVersions = $RemainingModuleVersions
  }
}

#endregion
