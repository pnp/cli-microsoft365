# teams tab add

Add a tab to the specified channel

## Usage

```sh
m365 teams tab add [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team to where the channel exists

`-c, --channelId <channelId>`
: The ID of the channel to add a tab to

`--appId <appId>`
: The ID of the Teams app that contains the Tab

`--appName <appName>`
: The name of the Teams app that contains the Tab

`--contentUrl <contentUrl>`
: The URL used for rendering Tab contents

`--entityId [entityId]`
: A unique identifier for the Tab

`--removeUrl [removeUrl]`
: The URL displayed when a Tab is removed

`--websiteUrl [websiteUrl]`
: The URL for showing tab contents outside of Teams

--8<-- "docs/cmd/_global.md"

## Remarks

The corresponding app must already be installed in the team.

## Examples
  
Add teams tab for website

```sh
m365 teams tab add --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --appId 06805b9e-77e3-4b93-ac81-525eb87513b8 --appName 'My Contoso Tab' --contentUrl 'https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView'
```

Add teams tab for website with additional configuration which is unknown

```sh
m365 teams tab add --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --appId 06805b9e-77e3-4b93-ac81-525eb87513b8 --appName 'My Contoso Tab' --contentUrl 'https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView' --test1 'value for test1'
```

## Response

=== "JSON"

    ``` json
    {
      "id": "8e454194-04c9-40aa-a9f3-7ab42d9541b5",
      "displayName": "'Polly'",
      "webUrl": "https://teams.microsoft.com/l/channel/19%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2/tab%3a%3a8e454194-04c9-40aa-a9f3-7ab42d9541b5?label=%27Polly%27&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022",
      "configuration": {
        "entityId": null,
        "contentUrl": "https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView",
        "removeUrl": null,
        "websiteUrl": null
      }
    }
    ```

=== "Text"

    ``` text
    configuration: {"entityId":null,"contentUrl":"https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView","removeUrl":null,"websiteUrl":null}
    displayName  : 'Polly'
    id           : 37d2294f-6dc0-4232-8718-d388f25ee696
    webUrl       : https://teams.microsoft.com/l/channel/19%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2/tab%3a%3a37d2294f-6dc0-4232-8718-d388f25ee696?label=%27Polly%27&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022
    ```

=== "CSV"

    ``` text
    id,displayName,webUrl,configuration
    0d7e343d-b233-4039-ae77-88928d4b275b,'Polly',https://teams.microsoft.com/l/channel/19%3aB3nCnLKwwCoGDEADyUgQ5kJ5Pkekujyjmwxp7uhQeAE1%40thread.tacv2/tab%3a%3a0d7e343d-b233-4039-ae77-88928d4b275b?label=%27Polly%27&groupId=aee5a2c9-b1df-45ac-9964-c708e760a045&tenantId=0cac6cda-2e04-4a3d-9c16-9c91470d7022,"{""entityId"":null,""contentUrl"":""https://www.contoso.com/Orders/2DCA2E6C7A10415CAF6B8AB6661B3154/tabView"",""removeUrl"":null,""websiteUrl"":null}"
    ```

