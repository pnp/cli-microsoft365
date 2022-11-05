# teams chat get

Get a Microsoft Teams chat conversation by id, participants or chat name.

## Usage

```sh
m365 teams chat get [options]
```

## Options

`-i, --id [id]`
: The ID of the chat conversation. Specify either `id`, `name` or `participants`, but not multiple.

`-n, --name [name]`
: The display name of the chat conversation. Specify either `id`, `name` or `participants`, but not multiple.

`-p, --participants [participants]`
: A comma-separated list of one or more e-mail addresses. Specify either `id`, `name` or `participants`, but not multiple.

--8<-- "docs/cmd/_global.md"

## Remarks

The output will not include the chat conversation members or messages. It will just retrieve the conversation details.
When using the `participants` option, the signed-in user will automatically be included as a participant. There's no need to add it to the list manually.

## Examples

Get a Microsoft Teams chat conversation by id

```sh
m365 teams chat get --id 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
```

Get a Microsoft Teams one on one chat conversation, finding it by participant.

```sh
m365 teams chat get --participants alexw@contoso.com
```

Get a Microsoft Teams group chat conversation, finding it by participants.

```sh
m365 teams chat get --participants alexw@contoso.com,meganb@contoso.com
```

Get a Microsoft Teams chat conversation, finding it by display name

```sh
m365 teams chat get --name "Just a conversation"
```

## Response

=== "JSON"

    ```json
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
    ```

=== "Text"

    ```text
    chatType           : oneOnOne
    createdDateTime    : 2022-11-05T13:06:25.218Z
    id                 : 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
    lastUpdatedDateTime: 2022-11-05T13:06:25.218Z
    onlineMeetingInfo  : null
    tenantId           : 446355e4-e7e3-43d5-82f8-d7ad8272d55b
    topic              : null
    viewpoint          : {"isHidden":false,"lastMessageReadDateTime":"2022-11-05T13:06:30.582Z"}
    webUrl             : https://teams.microsoft.com/l/chat/19%3A2da4c29f6d7041eca70b638b43d45437%40thread.v2/0?tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b
    ```

=== "CSV"

    ```csv
    id,topic,createdDateTime,lastUpdatedDateTime,chatType,webUrl,tenantId,onlineMeetingInfo,viewpoint
    19:2da4c29f6d7041eca70b638b43d45437@thread.v2,,2022-11-05T13:06:25.218Z,2022-11-05T13:06:25.218Z,oneOnOne,https://teams.microsoft.com/l/chat/19%3A2da4c29f6d7041eca70b638b43d45437%40thread.v2/0?tenantId=446355e4-e7e3-43d5-82f8-d7ad8272d55b,446355e4-e7e3-43d5-82f8-d7ad8272d55b,,"{""isHidden"":false,""lastMessageReadDateTime"":""2022-11-05T13:06:30.582Z""}"
    ```
