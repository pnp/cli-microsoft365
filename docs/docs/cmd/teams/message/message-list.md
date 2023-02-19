# teams message list

Lists all messages from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel for which to list messages

`-s, --since [since]`
: Date (ISO standard, dash separator) to get delta of messages from (in last 8 months)

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

## Examples

List the messages from a channel of the Microsoft Teams team

```sh
m365 teams message list --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype
```

List the messages from a channel of the Microsoft Teams team that have been created or modified since the date specified by the `--since` parameter (WARNING: only captures the last 8 months of data)

```sh
m365 teams message list --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype --since 2019-12-31T14:00:00Z
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "1666799217259",
        "replyToId": null,
        "etag": "1666799649208",
        "messageType": "message",
        "createdDateTime": "2022-10-26T15:46:57.259Z",
        "lastModifiedDateTime": "2022-10-26T15:54:09.208Z",
        "lastEditedDateTime": "2022-10-26T15:54:09.108Z",
        "deletedDateTime": null,
        "subject": "",
        "summary": null,
        "chatId": null,
        "importance": "normal",
        "locale": "en-us",
        "webUrl": "https://teams.microsoft.com/l/message/19%3eb30973b42a847a2a1df92d91e37c76a%40thread.tacv2/1666799217259?groupId=fce9e580-8bba-4638-ab5c-ab40016651e3&tenantId=92e59666-257b-49c3-b1fa-1bae8107f6ba&createdTime=1666799217259&parentMessageId=1666799217259",
        "policyViolation": null,
        "eventDetail": null,
        "from": {
          "application": null,
          "device": null,
          "user": {
            "id": "78ccf530-bbf0-47e4-aae6-da5f8c6fb142",
            "displayName": "John Doe",
            "userIdentityType": "aadUser",
            "tenantId": "92e59666-257b-49c3-b1fa-1bae8107f6ba"
          }
        },
        "body": {
          "contentType": "text",
          "content": "First message!"
        },
        "channelIdentity": {
          "teamId": "fce9e580-8bba-4638-ab5c-ab40016651e3",
          "channelId": "19:e2916df2b11046beba42d22da898383f@thread.tacv2"
        },
        "attachments": [],
        "mentions": [],
        "reactions": []
      }
    ]    
    ```

=== "Text"

    ``` text
    id             summary  body
    -------------  -------  ---------------
    1666799217259  null     First message!
    ```

=== "CSV"

    ``` text
    id,summary,body
    1666799217259,,First message!
    ```
