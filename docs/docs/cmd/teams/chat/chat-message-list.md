# teams chat message list

Lists all messages from a Microsoft Teams chat conversation.

## Usage

```sh
m365 teams chat message list [options]
```

## Options

`-i, --chatId <chatId>`
: The ID of the chat conversation

--8<-- "docs/cmd/_global.md"

## Examples

List the messages from a Microsoft Teams chat conversation

```sh
m365 teams chat message list --chatId 19:2da4c29f6d7041eca70b638b43d45437@thread.v2
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "1667653590582",
        "replyToId": null,
        "etag": "1667653590582",
        "messageType": "message",
        "createdDateTime": "2022-11-05T13:06:30.582Z",
        "lastModifiedDateTime": "2022-11-05T13:06:30.582Z",
        "lastEditedDateTime": null,
        "deletedDateTime": null,
        "subject": null,
        "summary": null,
        "chatId": "19:2da4c29f6d7041eca70b638b43d45437@thread.v2",
        "importance": "normal",
        "locale": "en-us",
        "webUrl": null,
        "channelIdentity": null,
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "John Doe",
            "userIdentityType": "aadUser",
            "tenantId": "446355e4-e7e3-43d5-82f8-d7ad8272d55b"
          }
        },
        "body": {
          "contentType": "html",
          "content": "<p>Hello world</p>"
        },
        "attachments": [],
        "mentions": [],
        "reactions": []
      }
    ]
    ```

=== "Text"

    ```text
    id             shortBody
    -------------  -------------------------
    1667653590582  <p>Hello world</p>
    ```

=== "CSV"

    ```csv
    id,shortBody
    1667653590582,<p>Hello world</p>
    ```

=== "Markdown"

    ```md
    # teams chat message list --chatId "19:04355ecd-2124-4097-bc2b-c2295a71d7a3_e1251b10-1ba4-49e3-b35a-933e3f21772b@unq.gbl.spaces"

    Date: 5/8/2023

    ## 1662642685689

    Property | Value
    ---------|-------
    id | 1662642685689
    etag | 1667653590582
    messageType | message
    createdDateTime | 2022-11-05T13:06:30.582Z
    lastModifiedDateTime | 2022-11-05T13:06:30.582Z
    chatId | 19:04355ecd-2124-4097-bc2b-c2295a71d7a3_e1251b10-1ba4-49e3-b35a-933e3f21772b@unq.gbl.spaces
    importance | normal
    locale | en-us
    body | <p>Hello world</p>
    shortBody | <p>Hello world</p>
    ```
