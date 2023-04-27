---
tags:
  - list
  - library
  - security
  - users
---

# Make a list or library read-only for all users

Author: [Nico De Cleyre](https://www.nicodecleyre.com), Inspired by [Salaudeen Rajack](https://www.sharepointdiary.com/2020/03/sharepoint-online-make-list-read-only-using-powershell.html)

The following script breaks the role inheritance of the given list if it doesn't have unique permissions already. After that it iterates through all role assignment and removes every role assignment that doesn't have the read permission. After that it gives the read permission to that assignment

=== "PowerShell"

    ```powershell
     param
    (
        [Parameter(Mandatory = $true, HelpMessage = "The site url where the list is located")][string] $SiteURL,
        [Parameter(Mandatory = $false, HelpMessage = "The name of the list")][string] $ListName
    )

    try {
      #Connect to m365
      $M365Status = m365 status
      if ($M365Status -match "Logged Out") {
        Write-Host "Logging in the User!"
        m365 login --authType browser
      }

      $List = m365 spo list get -u $SiteURL --title $ListName --withPermissions --properties "HasUniqueRoleAssignments" | ConvertFrom-Json

      #Break Permissions of the List
      if ($List.HasUniqueRoleAssignments -eq $False)
      {
        m365 spo list roleinheritance break --webUrl $SiteURL --listTitle $ListName --confirm
      }

      #Get Read Permission Level
      $RoleDefinitions =  m365 spo roledefinition list --webUrl $SiteURL | ConvertFrom-Json
      $ReadPermission = $RoleDefinitions | where {$_.RoleTypeKindValue -eq "Reader"}
        
      #Grant "Read" permissions, if its not granted already 
      ForEach($RoleAssignment in $List.RoleAssignments){
        if($RoleAssignment.Member.IsHiddenInUI -eq $False)
        {
          #Check if the current assignment has any permission other than Read or related
          $PermissionsToReplace = $RoleAssignment.RoleDefinitionBindings | Where {$_.Hidden -eq $False -And $_.Name -Notin ("Read", "Restricted Read", "Restricted Interfaces for Translation")}

          if($PermissionsToReplace -ne $Null)
          {
            #Remove All permissions
            ForEach($RoleDefBinding in $PermissionsToReplace)
            {
              m365 spo list roleassignment remove --webUrl $SiteURL --listTitle $ListName --principalId $RoleAssignment.Member.Id --confirm
              Write-Host "Removed all role assignments for '$($RoleDefBinding.Member.Title)'" -ForegroundColor Yellow
            }

            #Grant "Read" permissions, if it's not granted already
            m365 spo list roleassignment add --webUrl $SiteURL --listTitle $ListName --principalId $RoleAssignment.Member.Id --roleDefinitionId $ReadPermission.Id
            Write-Host "Added 'Read' Permissions to '$($RoleAssignment.Member.Title)'" -ForegroundColor Cyan
          }
        }
      }
      Write-host "List $($ListName) is set to Read-Only Successfully!" -f Green
    }
    catch {
        Write-Host -f Red "Error making the list $($ListName) read-only: " $_.Exception.Message
    }
    ```
