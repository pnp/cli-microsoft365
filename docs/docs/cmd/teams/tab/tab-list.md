# teams tab list

Lists tabs in the specified Microsoft Teams channel

## Usage

```sh
m365 teams tab list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel for which to list tabs

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve tabs for teams of which you are a member.

Tabs _Conversations_ and _Files_ are present in every team and therefore not included in the list of available tabs.

## Examples
  
List all tabs in a Microsoft Teams channel

```sh
m365 teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype
```

Include all the values from the tab configuration and associated teams app

```sh
m365 teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --output json
```

## Response

=== "JSON"

    ``` json
    [
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
        },
        "teamsApp": {
          "id": "1542629c-01b3-4a6d-8f76-1938b779e48d",
          "externalId": null,
          "displayName": "Polly",
          "distributionMethod": "store"
        },
        "teamsAppTabId": "1542629c-01b3-4a6d-8f76-1938b779e48d"
      }
    ]
    ```

=== "Text"

    ``` text
    displayName  : Polly
    id           : 34991fbf-59f4-48d9-b094-b9d64d550e23
    teamsAppTabId: 1542629c-01b3-4a6d-8f76-1938b779e48d
    ```

=== "CSV"

    ``` text
    id,displayName,teamsAppTabId
    34991fbf-59f4-48d9-b094-b9d64d550e23,Polly,1542629c-01b3-4a6d-8f76-1938b779e48d
    ```
