# code: insertSpaces=true tabSize=2

[CmdletBinding()]
param(
  [Parameter()]
  [String[]]$Name = @('*')
)

Set-StrictMode -Version Latest

# Upgrade installed modules
$UninstallModules = @(
  Get-InstalledModule -Name $Name |
  Sort-Object -Property Name |
  ForEach-Object {
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
      Write-Host -ForegroundColor Red "$($InstalledModule.Name) - Uninstalling unlisted version $($Module.Version) ($($Module.PublishedDate))"
      $Module | Uninstall-Module -Force -ErrorAction Continue
      return # Skip the rest of the ForEach-Object loop
    }
    # Upgrade module if needed
    if ($InstalledVersions.Version -NotContains $LatestModule.Version) {
      Write-Host -ForegroundColor Yellow "$($InstalledModule.Name) - Upgrading to $($LatestModule.Version) ($($LatestModule.PublishedDate))"
      Update-Module -Name $InstalledModule.Name -Force
    } else {
      Write-Host -ForegroundColor Green "$($InstalledModule.Name) - The latest version $($LatestModule.Version) ($($LatestModule.PublishedDate)) is already insalled"
    }
    # Uninstall old modules
    try {
      $InstalledVersions |
      Where-Object { $_.Version -ne $LatestModule.Version } |
      Sort-Object -Property Version |
      ForEach-Object {
        $Module = $_
        Write-Host -ForegroundColor Cyan "$($Module.Name) - Uninstalling old version $($Module.Version) ($($Module.PublishedDate))"
        $Module | Uninstall-Module -Force
      }
    } catch {
      # Catch modules with dependencies to be uninstalled later
      Get-InstalledModule -Name $InstalledModule.Name -AllVersions |
      Where-Object { $_.Version -ne $LatestModule.Version } |
      Sort-Object -Property Vesion
    }
  }
)

# Recursively uninstall old modules with dependencies
while ($UninstallModules) {
  $RetryModules = @()
  $UninstallModules |
  ForEach-Object {
    $Module = $_
    try {
      Write-Host -ForegroundColor Cyan "$($Module.Name) - Uninstalling old version $($Module.Version) ($($Module.PublishedDate))"
      $Module | Uninstall-Module -Force -ErrorAction Continue
    } catch {
      # Catch modules which still have dependencies to be uninstalled later
      $RetryModules += $Module
    }
  }
  $UninstallModules = $RetryModules
}
