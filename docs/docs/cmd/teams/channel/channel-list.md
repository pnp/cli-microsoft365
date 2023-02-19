# teams channel list

Lists channels in the specified Microsoft Teams team

## Usage

```sh
m365 teams channel list [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team to list the channels of. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team to list the channels of. Specify either `teamId` or `teamName` but not both

`--type [type]`
: Filter the results to only channels of a given type: `standard`, `private`, `shared`. By default all channels are listed.

--8<-- "docs/cmd/_global.md"

## Examples
  
List all channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000
```

List all channels in a specified Microsoft Teams team with name _Team Name_

```sh
m365 teams channel list --teamName "Team Name"
```

List private channels in a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000

```sh
m365 teams channel list --teamId 00000000-0000-0000-0000-000000000000 --type private
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2",
        "createdDateTime": "2022-10-26T15:43:31.954Z",
        "displayName": "Channel Name",
        "description": "This team is about Contoso",
        "isFavoriteByDefault": null,
        "email": "TeamsName@contoso.onmicrosoft.com",
        "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b",
        "webUrl": "https://teams.microsoft.com/l/channel/19%3AB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2/TeamsName?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b&allowXTenantAccess=False",
        "membershipType": "standard"
      }
    ]
    ```

=== "Text"

    ```text
    id                                                            displayName
    ------------------------------------------------------------  -----------
    19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2  Channel Name
    ```

=== "CSV"

    ```csv
    id,displayName
    19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2,Channel Name
    ```
