# Deploy Microsoft Teams personal app and add it to users

Author: [Sébastien Levert](https://github.com/sebastienlevert)

Installs or updates a Microsoft Teams app from a provided zipped manifest and then, based on the parameters, add it to the current users and / or to a set of users.

```powershell tab="PowerShell Core"
<#
  .SYNOPSIS
    Installs an app to Microsoft Teams and potentially to a set of users
  .EXAMPLE
    .\Install-PersonalApp.ps1 -AppManifestPath "C:\_\Manifest.zip" -AppId "2dbace6f-3f3b-4779-9e3f-bb4d27c403fe" -Publish -Install -CurrentUser
  .EXAMPLE
    .\Install-PersonalApp.ps1 -AppManifestPath "C:\_\Manifest.zip" -AppId "2dbace6f-3f3b-4779-9e3f-bb4d27c403fe" -Publish -Update -Install -Users @("user1@contoso.com", "user2@contoso.com")
  .PARAMETER AppManifestPath
    Valid Path to an zipped App Manifest representing the Microsoft Teams app
  .PARAMETER AppId
    GUID of the Microsoft Teams app. Is the same "id" you can find in the manifest.json from your Microsoft Teams app.
  .PARAMETER Users
    Array of string representing the usernames of the users to deploy the Microsoft Teams app to.
  .PARAMETER Publish
    Switch allowing to Publish (make available) the application to the tenant app catalog
  .PARAMETER Update
    Switch allowing to Update an existing application in the tenant app catalog
  .PARAMETER Install
    Switch allowing to Install the app for the specified Users or Current User
  .PARAMETER CurrentUser
    Switch allowing to Install the app for the current user
#>
Param(
  [ValidateScript( {
      if (-not ($_ | Test-Path) ) {
        throw "File or folder does not exist"
      }
      if (-not ($_ | Test-Path -PathType Leaf) ) {
        throw "The Path argument must be a file. Folder paths are not allowed."
      }
      if ($_ -notmatch ".zip") {
        throw "The file specified in the path argument must be a zip"
      }
      return $true
    })]
  [System.IO.FileInfo]$AppManifestPath,
  [string]$AppId,
  [string[]]$Users,
  [switch]$Publish,
  [switch]$Update,
  [switch]$Install,
  [switch]$CurrentUser
)

$m365Status = m365 status

if ($m365Status -eq "Logged Out") {
  # Connection to Microsoft 365
  m365 login
}

# Validating that the app is not already in the store
$app = m365 teams app list --query "[?externalId == '$AppId']" -o json | ConvertFrom-Json

if ($app.Length -gt 0) {
  if ($Update) {
    # Updating the app with the provided manifest
    m365 teams app update --id $app.id --filePath $AppManifestPath
    $app = m365 teams app list --query "[?externalId == '$AppId']" -o json | ConvertFrom-Json
    Write-Host "The App '$($app.displayName)' with ID '$($app.id)' and ExternalID '$($app.externalId)' was updated."
  }
}
else {
  if ($Publish) {
    # Publishing the app with the provided manifest
    m365 teams app publish --filePath $AppManifestPath
    $app = m365 teams app list --query "[?externalId == '$AppId']" -o json | ConvertFrom-Json
    Write-Host "The App '$($app.displayName)' with ID '$($app.id)' and ExternalID '$($app.externalId)' was published."
  }
}

if ($CurrentUser) {
  # Getting the reference of the currently connected user
  $connectedAs = m365 status -o json | ConvertFrom-Json
  $user = m365 aad user get --userName $connectedAs.connectedAs -o json | ConvertFrom-Json

  if ($user) {
    $Users += $user.userPrincipalName
  }
}  

$user = $null
if ($Users.Length -gt 0 -and $Install) {
  $Users | ForEach-Object {
    # Getting the specified user
    $user = m365 aad user get --userName $_ -o json | ConvertFrom-Json
  
    if ($user) {
      $userApp = m365 teams user app list --userId $user.id --query "[?appId == '$($app.id)']" -o json | ConvertFrom-Json

      if ($userApp.Length -eq 0) {
        # Adding the app to the personal apps of the specified user
        m365 teams user app add --appId $app.id --userId $user.id
        Write-Host "The App '$($app.displayName)' with ID '$($app.id)' was deployed to user '$($user.userPrincipalName)'."

      }
      else {
        Write-Warning "The App '$($app.displayName)' with ID '$($app.id)' is already deployed to user '$($user.userPrincipalName)'."
      }
    }
    else {
      Write-Warning "The user '$_' was not found"
    }
  }
}
```

Keywords:

- Microsoft Teams
