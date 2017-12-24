$date = [DateTime]::UtcNow

#Update main assembly version number in AssemblyInfo.cs
$year = $date | Get-Date -format yy
$dayNumber = $date.DayOfYear
$build = $year + $dayNumber
$revision = $date | Get-Date -format HHmm

Write-Host "New Build:" $build
Write-Host "New Revision:" $revision

$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$project_path = Split-Path $script_folder -Parent
$assemblyinfo_path = Join-Path $project_path "CrmDeveloperExtensions2/Properties"
$assemblyinfo = Join-Path $assemblyinfo_path "AssemblyInfo.cs"

$assemblyinfo_content = [System.IO.File]::ReadAllText($assemblyinfo)

$assemblyVersionPattern = 'AssemblyVersion\("[0-9]+(\.([0-9]+|\*)){1,3}"\)'
$versionString = [RegEx]::Match($assemblyinfo_content, $assemblyVersionPattern)
$versionName = [RegEx]::Match($versionString, "((?:\d+\.\d+\.\d+\.\d+))")
$version = [version] $versionName.Value

Write-Host "Current Assembly Major: " $version.Major
Write-Host "Current Assembly Minor: " $version.Minor
Write-Host "Current Assembly Build: " $version.Build
Write-Host "Current Assembly Revision: " $version.Revision

$newVersion = $version.Major.ToString() + "." + $version.Minor.ToString() + "." + $build + "." + $revision
Write-Host "New Version: " $newVersion

$assemblyinfo_content = [Regex]::Replace($assemblyinfo_content, "(\d+)\.(\d+)\.(\d+)[\.(\d+)]*", $newVersion)

#Update main assembly copyright in AssemblyInfo.cs
$copyrightPattern = 'AssemblyCopyright\(".*"\)\]'
$newCopyright = 'Copyright ©  ' + $date.Year
Write-Host ("New Assembly Copyright: " + $newCopyright)

$assemblyinfo_content = [Regex]::Replace($assemblyinfo_content, $copyrightPattern, 'AssemblyCopyright("' + $newCopyright + '")]')

[System.IO.File]::WriteAllText($assemblyinfo, $assemblyinfo_content)

#Update manifest version in source.extension.vsixmanifest
$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$project_path = Split-Path $script_folder -Parent
$manifest_path = Join-Path $project_path "CrmDeveloperExtensions2"
$manifest = Join-Path $manifest_path "source.extension.vsixmanifest"

$manifest_content = [System.IO.File]::ReadAllText($manifest)

$manifestPattern = 'Version="(\d+)\.(\d+)\.(\d+).(\d+)" Language='
$manifestVersionString = [RegEx]::Match($manifest_content, $manifestPattern)
$manifestVersionName = [RegEx]::Match($manifestVersionString, "(\d+)\.(\d+)\.(\d+).(\d+)")
$manifestVersion = [version] $manifestVersionName.Value

Write-Host "Current Manifest Major: " $manifestVersion.Major
Write-Host "Current Manifest Minor: " $manifestVersion.Minor
Write-Host "Current Manifest Build: " $manifestVersion.Build
Write-Host "Current Manifest Revision: " $manifestVersion.Revision

$manifest_content = [Regex]::Replace($manifest_content, $manifestPattern, 'Version="' + $newVersion + '" Language=')

[System.IO.File]::WriteAllText($manifest, $manifest_content)

#Update package version in CrmDeveloperExtensions2Package.cs
$script_path = $myinvocation.mycommand.path
$script_folder = Split-Path $script_path -Parent
$project_path = Split-Path $script_folder -Parent
$package_path = Join-Path $project_path "CrmDeveloperExtensions2"
$package = Join-Path $package_path "CrmDeveloperExtensions2Package.cs"

$package_content = [System.IO.File]::ReadAllText($package)

$packagePattern = '(\d+)\.(\d+)\.(\d+).(\d+)'
$packageVersionString = [RegEx]::Match($package_content, $packagePattern)
$packageVersionName = [RegEx]::Match($packageVersionString, "(\d+)\.(\d+)\.(\d+).(\d+)")
$packageVersion = [version] $packageVersionName.Value

Write-Host "Current Package Major: " $packageVersion.Major
Write-Host "Current Package Minor: " $packageVersion.Minor
Write-Host "Current Package Build: " $packageVersion.Build
Write-Host "Current Package Revision: " $packageVersion.Revision

$package_content = [Regex]::Replace($package_content, $packagePattern, $newVersion)

[System.IO.File]::WriteAllText($package, $package_content)

Write-Host ("##vso[task.setvariable variable=BuildVersion;]$newVersion")