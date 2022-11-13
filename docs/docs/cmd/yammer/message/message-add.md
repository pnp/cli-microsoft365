# yammer message add

Posts a Yammer network message on behalf of the current user

## Usage

```sh
m365 yammer message add [options]
```

## Options

`-b, --body <body>`
: The text of the message body

`--groupId [groupId]`
: Post the message to this group, specified by ID. If this is set then the networkId is inferred from it. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`-r, --repliedToId [repliedToId]`
: The message ID this message is in reply to. If this is set then groupId and networkId are inferred from it. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`-d, --directToUserIds [directToUserIds]`
: Send a private message to one or more users, specified by ID. Alternatively, you can use the Yammer network e-mail addresses instead of the IDs. You must either specify `groupId`, `repliedToId`, or `directToUserIds` to send the message

`--networkId [networkId]`
: Specify the network to post a message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Replies to a message with the ID 1231231231

```sh
m365 yammer message add --body "Hello everyone!" --repliedToId 1231231231
```

Sends a private conversation to the user with the ID 1231231231

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds 1231231231
```

Sends a private conversation to multiple users by ID

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds "1231231231,1121312"
```

Sends a private conversation to the user with the e-mail pl@nubo.eu and sc@nubo.eu

```sh
m365 yammer message add --body "Hello everyone!" --directToUserIds "pl@nubo.eu,sc@nubo.eu"
```

Posts a message to the group with the ID 12312312312

```sh
m365 yammer message add --body "Hello everyone!" --groupId 12312312312
```

Posts a message to the group with the ID 12312312312 in the network 11112312

```sh
m365 yammer message add --body "Hello everyone!" --groupId 12312312312 --networkId 11112312
```

## Response

=== "JSON"

    ```json
    {
      "id": 2000337346863105,
      "sender_id": 36425097217,
      "delegate_id": null,
      "replied_to_id": null,
      "created_at": "2022/11/11 20:59:56 +0000",
      "network_id": 5897756673,
      "message_type": "update",
      "sender_type": "user",
      "url": "https://www.yammer.com/api/v1/messages/2000337346863105",
      "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/messages/2000337346863105",
      "group_id": 31158067201,
      "body": {
        "parsed": "Hello everyone!",
        "plain": "Hello everyone!",
        "rich": "Hello everyone!"
      },
      "thread_id": 2000337346863105,
      "client_type": "O365 Api Auth",
      "client_url": "https://api.yammer.com",
      "system_message": false,
      "direct_message": false,
      "chat_client_sequence": null,
      "language": null,
      "notified_user_ids": [],
      "privacy": "public",
      "attachments": [],
      "liked_by": {
        "count": 0,
        "names": []
      },
      "supplemental_reply": false,
      "content_excerpt": "Hello everyone!",
      "group_created_id": 31158067201
    }
    ```

=== "Text"

    ```text
    id: 2000337648877569
    ```

=== "CSV"

    ```csv
    id
    2000337749565441
    ```
