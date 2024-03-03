npm login

Write-Host "Deprecate beta versions of the @pnp/cli-microsoft365 npm package on npm"
$version = Read-Host "Version of the package to deprecate"
$otp = Read-Host "One-time password"
$allVersions = npm view @pnp/cli-microsoft365 versions --json | ConvertFrom-Json
$versionsToDeprecate = $allVersions | Where-Object { $_ -ne $null -and $_.StartsWith("$version-beta.") }

if ($versionsToDeprecate.Length -eq 0) {
  Write-Host "No versions matching $version-beta found"
  return
}

Write-Host "Deprecating $($versionsToDeprecate.Length) versions of the package"

$index = 0
$total = $versionsToDeprecate.Length

$versionsToDeprecate | ForEach-Object {
  $index = $index + 1

  Write-Host "$index/$total - Processing deprecation '$_'..." -ForegroundColor Green

  $versionInfo = npm view "@pnp/cli-microsoft365@$_" --json | ConvertFrom-Json

  if ($null -eq $versionInfo.deprecated) {
    Write-Host "Deprecating..."
    npm deprecate "@pnp/cli-microsoft365@$_" "Preview version released" --otp=$otp
  } else {
    Write-Host "Version already deprecated..."
  }
}