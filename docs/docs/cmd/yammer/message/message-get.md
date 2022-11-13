# yammer message get

Returns a Yammer message

## Usage

```sh
m365 yammer message get [options]
```

## Options

`--id <id>`
: The id of the Yammer message

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

## Examples

Returns the Yammer message with the id 1239871123

```sh
m365 yammer message get --id 1239871123
```

Returns the Yammer message with the id 1239871123 in JSON format

```sh
m365 yammer message get --id 1239871123 --output json
```

## Response

=== "JSON"

    ```json
    {
      "id": 2000337749565441,
      "sender_id": 36425097217,
      "delegate_id": null,
      "replied_to_id": null,
      "created_at": "2022/11/11 21:00:20 +0000",
      "network_id": 5897756673,
      "message_type": "update",
      "sender_type": "user",
      "url": "https://www.yammer.com/api/v1/messages/2000337749565441",
      "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/messages/2000337749565441",
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
      "content_excerpt": "Hello everyone!"
    }
    ```

=== "Text"

    ```text
    content_excerpt: Hello everyone!
    created_at     : 2022/11/11 21:00:20 +0000
    direct_message : false
    group_id       : 31158067201
    id             : 2000337749565441
    message_type   : update
    privacy        : public
    replied_to_id  : null
    sender_id      : 36425097217
    system_message : false
    thread_id      : 2000337749565441
    ```

=== "CSV"

    ```csv
    id,sender_id,replied_to_id,thread_id,group_id,created_at,direct_message,system_message,privacy,message_type,content_excerpt
    2000337749565441,36425097217,,2000337749565441,31158067201,2022/11/11 21:00:20 +0000,,,public,update,Hello everyone!
    ```
