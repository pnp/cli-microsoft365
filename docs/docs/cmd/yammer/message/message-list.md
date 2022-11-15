# yammer message list

Returns all accessible messages from the user's Yammer network

## Usage

```sh
m365 yammer message list [options]
```

## Options

`--olderThanId [olderThanId]`
: Returns messages older than the message ID specified as a numeric string

`--threaded`
: Will only return the thread starter (first message) for each thread. This parameter is intended for apps which need to display message threads collapsed

`-f, --feedType [feedType]`
: Returns messages from a specific feed. Available options: `All,Top,My,Following,Sent,Private,Received`. Default `All`

`--groupId [groupId]`
: Returns the messages from a specific group

`--threadId [threadId]`
: Returns the messages from a specific thread

`--limit [limit]`
: Limits the messages returned

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

Feed types

- All: Corresponds to “All” conversations in the Yammer web interface
- Top: The algorithmic feed for the user that corresponds to "Top" conversations. The Top conversations feed is the feed currently shown in the Yammer mobile apps
- My: The user’s feed, based on the selection they have made between "Following" and "Top" conversations
- Following: The "Following" feed which is conversations involving people and topics that the user is following
- Sent: All messages sent by the user
- Private: Private messages received by the user
- Received: All messages received by the user

## Examples

Returns all Yammer network messages

```sh
m365 yammer message list
```

Returns all Yammer network messages older than the message ID 5611239081

```sh
m365 yammer message list --olderThanId 5611239081
```

Returns all Yammer network thread starter (first message) for each thread

```sh
m365 yammer message list --threaded
```

Returns the first 10 Yammer network messages

```sh
m365 yammer message list --limit 10
```

Returns the first 10 Yammer network messages from the Yammer group 312891231

```sh
m365 yammer message list --groupId 312891231 --limit 10
```

Returns the first 10 Yammer network messages from thread 5611239081

```sh
m365 yammer message list --threadId 5611239081 --limit 10
```

Returns the first 20 Yammer message from the sent feed of the user

```sh
m365 yammer message list --feedType Sent --limit 20
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": 2000337749565441,
        "sender_id": 36425097217,
        "delegate_id": null,
        "replied_to_id": null,
        "created_at": "2022/11/11 20:59:10 +0000",
        "network_id": 5897756673,
        "message_type": "update",
        "sender_type": "user",
        "url": "https://www.yammer.com/api/v1/messages/2000336575053825",
        "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/messages/2000336575053825",
        "group_id": 31158067201,
        "body": {
          "parsed": "Hello everyone!",
          "plain": "Hello everyone!",
          "rich": "Hello everyone!"
        },
        "thread_id": 2000337749565441,
        "client_type": "O365 Api Auth",
        "client_url": "https://api.yammer.com",
        "system_message": false,
        "direct_message": false,
        "chat_client_sequence": null,
        "language": "no",
        "notified_user_ids": [],
        "privacy": "public",
        "attachments": [],
        "liked_by": {
          "count": 0,
          "names": []
        },
        "supplemental_reply": false,
        "content_excerpt": "Hello everyone!",
        "group_created_id": 31158067201,
        "shortBody": "Hello everyone!"
      }
    ]
    ```

=== "Text"

    ```text
    id                replied_to_id  thread_id         group_id     shortBody
    ----------------  -------------  ----------------  -----------  ---------------
    2000337749565441  null           2000337749565441  31158067201  Hello everyone!
    ```

=== "CSV"

    ```csv
    id,replied_to_id,thread_id,group_id,shortBody
    2000337749565441,,2000337749565441,31158067201,Hello everyone!
    ```
