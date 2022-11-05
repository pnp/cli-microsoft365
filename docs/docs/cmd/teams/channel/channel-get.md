# teams channel get

Gets information about the specific Microsoft Teams team channel

## Usage

```sh
m365 teams channel get [options]
```

## Options

`--teamId [teamId]`
: The ID of the team to which the channel belongs to. Specify either `teamId` or `teamName` but not both.

`--teamName [teamName]`
: The display name of the team to which the channel belongs to. Specify either `teamId` or `teamName` but not both.

`-i, --id [id]`
: The ID of the channel for which to retrieve more information. Specify either `id` or `name` but not both.

`--name [name]`
: The display name of the channel for which to retrieve more information. Specify either `id` or `name` but not both.

`--primary`
: Gets the default channel, General, of a team. If specified, id or name are not needed.

--8<-- "docs/cmd/_global.md"

## Examples
  
Get information about Microsoft Teams team channel with id _19:493665404ebd4a18adb8a980a31b4986@thread.skype_

```sh
m365 teams channel get --teamId 00000000-0000-0000-0000-000000000000 --id '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
```

Get information about Microsoft Teams team channel with name _Channel Name_

```sh
m365 teams channel get --teamName "Team Name" --name "Channel Name"
```

Get information about Microsoft Teams team primary channel , i.e. General

```sh
m365 teams channel get --teamName "Team Name" --primary
```

## Response

=== "JSON"

    ```json
    {
      "id": "19:493665404ebd4a18adb8a980a31b4986@thread.tacv2",
      "createdDateTime": "2022-10-26T15:43:31.954Z",
      "displayName": "Channel Name",
      "description": "This team is about Contoso",
      "isFavoriteByDefault": null,
      "email": "TeamName@contoso.onmicrosoft.com",
      "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3A493665404ebd4a18adb8a980a31b4986%40thread.tacv2/ChannelName?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b&allowXTenantAccess=False",
      "membershipType": "standard"
    }
    ```

=== "Text"

    ```text
    createdDateTime    : 2022-10-26T15:43:31.954Z
    description        : This team is about the Contoso
    displayName        : Channel Name
    email              : TeamName@ordidev.onmicrosoft.com
    id                 : 19:493665404ebd4a18adb8a980a31b4986@thread.tacv2
    isFavoriteByDefault: null
    membershipType     : standard
    tenantId           : 446355e4-e7e3-43d5-82f8-d7ad8272d55b
    webUrl             : https://teams.microsoft.com/l/channel/19%3A493665404ebd4a18adb8a980a31b4986%40thread.tacv2/ChannelName?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b&allowXTenantAccess=False
    ```

=== "CSV"

    ```csv
    id,createdDateTime,displayName,description,isFavoriteByDefault,email,tenantId,webUrl,membershipType
    19:493665404ebd4a18adb8a980a31b4986@thread.tacv2,2022-10-26T15:43:31.954Z,Channel Name,This team is about Contoso,,TeamName@contoso.onmicrosoft.com,446355e4-e7e3-43d5-82f8-d7ad8272d55b,https://teams.microsoft.com/l/channel/19%3A493665404ebd4a18adb8a980a31b4986%40thread.tacv2/ChannelName?groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b&allowXTenantAccess=False,standard
    ```
