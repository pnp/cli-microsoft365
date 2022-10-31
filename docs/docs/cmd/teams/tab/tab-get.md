# teams tab get

Gets information about the specified Microsoft Teams tab

## Usage

```sh
m365 teams tab get [options]
```

## Options

`--teamId [teamId]`
: The ID of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both

`--teamName [teamName]`
: The display name of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both

`--channelId [channelId]`
: The ID of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both

`--channelName [channelName]`
: The display name of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both

`-i, --tabId [tabId]`
: The ID of the Microsoft Teams tab. Specify either tabId or tabName but not both

`-n, --tabName [tabName]`
: The display name of the Microsoft Teams tab. Specify either tabId or tabName but not both

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve tabs for teams of which you are a member.

## Examples

Get a Microsoft Teams Tab with ID _1432c9da-8b9c-4602-9248-e0800f3e3f07_

```sh
m365 teams tab get --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07
```

Get a Microsoft Teams Tab with name _Tab Name_

```sh
m365 teams tab get --teamName "Team Name" --channelName "Channel Name" --tabName "Tab Name"
```

## Response

=== "JSON"

    ``` json
    {
      "id": "34991fbf-59f4-48d9-b094-b9d64d550e23",
      "displayName": "Polly",
      "webUrl": "https://teams.microsoft.com/l/entity/1542629c-01b3-4a6d-8f76-1938b779e48d/_djb2_msteams_prefix_34991fbf-59f4-48d9-b094-b9d64d550e23?webUrl=https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d&label=Polly&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d%22%2c%0d%0a++%22channelId%22%3a+%2219%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022",
      "configuration": {
        "entityId": "surveys_list:19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2:ps67c9jyf3a30j2j5eum72",
        "contentUrl": "https://teams.polly.ai/msteams/content/tab/team?theme={theme}",
        "removeUrl": "https://teams.polly.ai/msteams/content/tabdelete?theme={theme}",
        "websiteUrl": "https://teams.polly.ai/msteams/content/tab/team?theme={theme}",
        "dateAdded": "2022-10-31T12:17:58.632Z"
      }
    }
  ```

=== "Text"

    ``` text
    configuration: {"entityId":"surveys_list:19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2:ps67c9jyf3a30j2j5eum72","contentUrl":"https://teams.polly.ai/msteams/content/tab/team?theme={theme}","removeUrl":"https://teams.polly.ai/msteams/content/tabdelete?theme={theme}","websiteUrl":"https://teams.polly.ai/msteams/content/tab/team?theme={theme}","dateAdded":"2022-10-31T12:17:58.632Z"}
    displayName  : Polly
    id           : 34991fbf-59f4-48d9-b094-b9d64d550e23
    webUrl       : https://teams.microsoft.com/l/entity/1542629c-01b3-4a6d-8f76-1938b779e48d/_djb2_msteams_prefix_34991fbf-59f4-48d9-b094-b9d64d550e23?webUrl=https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d&label=Polly&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d%22%2c%0d%0a++%22channelId%22%3a+%2219%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022
    ```

=== "CSV"

    ``` text
    id,displayName,webUrl,configuration
    34991fbf-59f4-48d9-b094-b9d64d550e23,Polly,https://teams.microsoft.com/l/entity/1542629c-01b3-4a6d-8f76-1938b779e48d/_djb2_msteams_prefix_34991fbf-59f4-48d9-b094-b9d64d550e23?webUrl=https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d&label=Polly&context=%7b%0d%0a++%22canvasUrl%22%3a+%22https%3a%2f%2fteams.polly.ai%2fmsteams%2fcontent%2ftab%2fteam%3ftheme%3d%7btheme%7d%22%2c%0d%0a++%22channelId%22%3a+%2219%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2%22%2c%0d%0a++%22subEntityId%22%3a+null%0d%0a%7d&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022,"{""entityId"":""surveys_list:19:B3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1@thread.tacv2:ps67c9jyf3a30j2j5eum72"",""contentUrl"":""https://teams.polly.ai/msteams/content/tab/team?theme={theme}"",""removeUrl"":""https://teams.polly.ai/msteams/content/tabdelete?theme={theme}"",""websiteUrl"":""https://teams.polly.ai/msteams/content/tab/team?theme={theme}"",""dateAdded"":""2022-10-31T12:17:58.632Z""}"
    ```
