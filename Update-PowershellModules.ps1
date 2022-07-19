# code: insertSpaces=true tabSize=2

[CmdletBinding()]
param(
  [Parameter()]
  [String[]]$Name = @('*')
)

Set-StrictMode -Version Latest

# Upgrade installed modules
Get-InstalledModule -Name $Name |
Sort-Object -Property Name |
ForEach-Object -Begin {
  # Obsolete modules with dependencies to be uninstalled later
  [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '')]
  $ObsoleteModuleVersions = @()
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
    Write-Host -ForegroundColor Green "$($InstalledModule.Name) - The latest version $($LatestModule.Version) ($($LatestModule.PublishedDate)) is already insalled"
  }
  # Uninstall obsolete module versions
  try {
    $InstalledVersions |
    Where-Object { $_.Version -ne $LatestModule.Version } |
    Sort-Object -Property Version |
    ForEach-Object {
      $ObsoleteModuleVersion = $_
      Write-Host -ForegroundColor Cyan "$($ObsoleteModuleVersion.Name) - Uninstalling obsolete version $($ObsoleteModuleVersion.Version) ($($ObsoleteModuleVersion.PublishedDate))"
      $ObsoleteModuleVersion | Uninstall-Module -Force
    }
  } catch {
    # Catch obsolete module version with dependencies to be uninstalled later
    $ObsoleteModuleVersions += $ObsoleteModuleVersion
  }
} -End {
  # Recursively uninstall obsolete module versions with dependencies
  while ($ObsoleteModuleVersions) {
    $RemainingModuleVersions = @()
    $ObsoleteModuleVersions |
    ForEach-Object {
      $ObsoleteModuleVersion = $_
      try {
        Write-Host -ForegroundColor Cyan "$($ObsoleteModuleVersion.Name) - Uninstalling obsolete version $($ObsoleteModuleVersion.Version) ($($ObsoleteModuleVersion.PublishedDate))"
        $ObsoleteModuleVersion | Uninstall-Module -Force -ErrorAction Continue
      } catch {
        # Catch module versions which still have dependencies to be uninstalled later
        $RemainingModuleVersions += $ObsoleteModuleVersion
      }
    }
    # No module versions were uninstalled due to dependencies on modules outside the updated module name scope
    if ($RemainingModuleVersions.Count -eq $ObsoleteModuleVersions.Count) {
      $RemainingModuleVersions |
      ForEach-Object {
        $ObsoleteModuleVersion = $_
        Write-Host -ForegroundColor Red "$($ObsoleteModuleVersion.Name) - Cannot uninstall obsolete version $($ObsoleteModuleVersion.Version) ($($ObsoleteModuleVersion.PublishedDate)) because of dependencies!"
      }
      # Stop uninstalling modules with dependencies on modules outside the updated module name scope
      $RemainingModuleVersions = @()
    }
    # Uninstall the remaining modules
    $ObsoleteModuleVersions = $RemainingModuleVersions
  }
}
