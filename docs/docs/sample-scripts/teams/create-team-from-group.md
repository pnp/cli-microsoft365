# Create a Microsoft Team with channels from a Microsoft 365 Group

Inspired by: [Patrick Lamber](https://www.nubo.eu/Provision-A-Team-With-CLI-For-Microsoft-365/)

A sample script which creates a Microsoft 365 Group, associates a logo to it and some members. Afterward, it teamyfies the Group and creates two public channels.

=== "PowerShell"

    ```powershell
    # this examples searches the users in a directory by displayname
    $memberDisplayName = "A"
    # Group settings
    $logoPath = "./pnpImage.png"
    $displayName = "Contoso Group"
    $mailNickName = "contosoGroup"
    # add more items to the array to provision channels
    $channelNames = @("Public relations", "CLI Project")

    Write-Host "Creating the Group '$displayName'..."
    $group = $null
    $group = m365 aad o365group add --displayName $displayName `
                                    --description "." --mailNickname $mailNickName  `
                                    -o "json" | convertfrom-json
    if ($group -eq $null) {
        Write-Host "An error occurred during Group creation"
        break
    }

    Write-Host "Created with id $($group.id)"

    # you might need to wait a little bit after Group creation before you are allowed to assign a logo
    Write-Host "Assigning custom logo '$logoPath' in about 10 seconds..."
    Start-sleep -Seconds 10
    m365 aad o365group set --id $group.id --logoPath $logoPath    

    Write-Host "Searching for members with '$memberDisplayName' in their displayname"
    $membersToAdd = m365 aad user list --displayName $memberDisplayName --properties "id,userprincipalname" --output "json" | convertfrom-json
    $membersToAdd | ForEach-Object {
        Write-Host "Adding member to $($_.userPrincipalName) to Group"
        $variable = m365 aad o365group user add --groupId $group.id --userName $_.id -o "json" | convertfrom-json
    }

    Write-Host "Teamify the Group"
    m365 aad o365group teamify --groupId $group.id

    Write-Host "Provisioning channels"
    $channelNames | ForEach-Object {
        m365 teams channel add --teamId $group.id --name $_ 
    }
    ```

Keywords:

- Microsoft Teams
- Provisioning
