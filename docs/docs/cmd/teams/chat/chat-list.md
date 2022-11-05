# teams chat list

Lists all Microsoft Teams chat conversations for the current user.

## Usage

```sh
m365 teams chat list [options]
```

## Options

`-t, --type [chatType]`
: The chat type to optionally filter chat conversations by type. The value can be `oneOnOne`, `group` or `meeting`.

--8<-- "docs/cmd/_global.md"

## Examples

List all the Microsoft Teams chat conversations of the current user.

```sh
m365 teams chat list
```

List only the one on one Microsoft Teams chat conversations.

```sh
m365 teams chat list --type oneOnOne
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
        "topic": null,
        "createdDateTime": "2022-11-05T13:06:25.218Z",
        "lastUpdatedDateTime": "2022-11-05T13:06:25.218Z",
        "chatType": "oneOnOne",
        "webUrl": "https://teams.microsoft.com/l/chat/19%3A2da4c29f6d7041eca70b638b43d45437%40thread.v2/0?tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b",
        "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b",
        "onlineMeetingInfo": null,
        "viewpoint": {
          "isHidden": false,
          "lastMessageReadDateTime": "2022-11-05T13:06:30.582Z"
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                             topic  chatType
    ---------------------------------------------  -----  --------
    19:2da4c29f6d7041eca70b638b43d45437@thread.v2  null   oneOnOne
    ```

=== "CSV"

    ```csv
    id,topic,chatType
    19:2da4c29f6d7041eca70b638b43d45437@thread.v2,,oneOnOne
    ```
