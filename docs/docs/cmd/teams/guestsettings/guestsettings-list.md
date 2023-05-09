# teams guestsettings list

Lists guest settings for a Microsoft Teams team

## Usage

```sh
m365 teams guestsettings list [options]
```

## Options

`-i, --teamId`
: The ID of the team for which to get the guest settings

--8<-- "docs/cmd/_global.md"

## Examples

Get guest settings for a Microsoft Teams team

```sh
m365 teams guestsettings list --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```

## Response

=== "JSON"

    ```json
    {
      "allowCreateUpdateChannels": false,
      "allowDeleteChannels": false
    }
    ```

=== "Text"

    ```text
    allowCreateUpdateChannels: false
    allowDeleteChannels      : false
    ```

=== "CSV"

    ```csv
    allowCreateUpdateChannels,allowDeleteChannels
    ,
    ```

==="Markdown"

    ```md
    # teams guestsettings list --teamId "2609af39-7775-4f94-a3dc-0dd67657e900"

    Date: 5/7/2023

    Property | Value
    ---------|-------
    allowCreateUpdateChannels | false
    allowDeleteChannels | false
   ```
