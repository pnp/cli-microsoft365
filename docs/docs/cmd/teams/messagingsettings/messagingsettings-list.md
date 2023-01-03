# teams messagingsettings list

Lists messaging settings for a Microsoft Teams team

## Usage

```sh
m365 teams messagingsettings list [options]
```

## Options

`-i, --teamId`
: The ID of the team for which to get the messaging settings.

--8<-- "docs/cmd/_global.md"

## Examples

Get messaging settings for a Microsoft Teams team.

```sh
m365 teams messagingsettings list --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```

## Response

=== "JSON"

    ``` json
    {
      "allowUserEditMessages": true,
      "allowUserDeleteMessages": true,
      "allowOwnerDeleteMessages": true,
      "allowTeamMentions": true,
      "allowChannelMentions": true
    }
    ```

=== "Text"

    ``` text
    allowChannelMentions    : true
    allowOwnerDeleteMessages: true
    allowTeamMentions       : true
    allowUserDeleteMessages : true
    allowUserEditMessages   : true
    ```

=== "CSV"

    ``` text
    allowUserEditMessages,allowUserDeleteMessages,allowOwnerDeleteMessages,allowTeamMentions,allowChannelMentions
    1,1,1,1,1
    ```

=== "Markdown"

    ```md
    # teams messagingsettings list --teamId "2609af39-7775-4f94-a3dc-0dd67657e900"

    Date: 1/3/2023

    ## undefined (undefined)

    Property | Value
    ---------|-------
    allowUserEditMessages | true
    allowUserDeleteMessages | true
    allowOwnerDeleteMessages | true
    allowTeamMentions | true
    allowChannelMentions | true
    ```
