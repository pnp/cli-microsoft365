# teams membersettings list

Lists member settings for a Microsoft Teams team

## Usage

```sh
m365 teams membersettings list [options]
```

## Options

`-i, --teamId`
: The ID of the team for which to get the member settings

--8<-- "docs/cmd/_global.md"

## Examples

Get member settings for a Microsoft Teams team

```sh
m365 teams membersettings list --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```

## Response

=== "JSON"

    ```json
    {
      "allowCreateUpdateChannels": true,
      "allowCreatePrivateChannels": true,
      "allowDeleteChannels": true,
      "allowAddRemoveApps": true,
      "allowCreateUpdateRemoveTabs": true,
      "allowCreateUpdateRemoveConnectors": true
    }
    ```

=== "Text"

    ```text
    allowAddRemoveApps               : true
    allowCreatePrivateChannels       : true
    allowCreateUpdateChannels        : true
    allowCreateUpdateRemoveConnectors: true
    allowCreateUpdateRemoveTabs      : true
    allowDeleteChannels              : true
    ```

=== "CSV"

    ```csv
    allowCreateUpdateChannels,allowCreatePrivateChannels,allowDeleteChannels,allowAddRemoveApps,allowCreateUpdateRemoveTabs,allowCreateUpdateRemoveConnectors
    1,1,1,1,1,1
    ```

=== "Markdown"

    ```md
    # teams membersettings list --teamId "2609af39-7775-4f94-a3dc-0dd67657e900"

      Date: 5/7/2023

      Property | Value
      ---------|-------
      allowCreateUpdateChannels | true
      allowCreatePrivateChannels | true
      allowDeleteChannels | true
      allowAddRemoveApps | true
      allowCreateUpdateRemoveTabs | true
      allowCreateUpdateRemoveConnectors | true
     ```
