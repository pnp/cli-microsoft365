# yammer search

Returns a list of messages, users, topics and groups that match the specified query.

## Usage

```sh
m365 yammer search [options]
```

## Options

`--queryText <queryText>`
: The query for the search

`--show [show]`
: Specifies the type of data to return when using --output text. Allowed values `summary, messages, users, topics, groups`.

`--limit [limit]`
: Limits the results returned for each item category.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    In order to use this command, you need to grant the Azure AD application used by the CLI for Microsoft 365 the permission to the Yammer API. To do this, execute the `cli consent --service yammer` command.

Using the `--show` option in JSON output is not supported. To filter JSON results, use either a JMESPath query or filter the data yourself after retrieving them.

## Examples

Returns search result for the query `community`

```sh
m365 yammer search --queryText "community"
```

Returns groups that match `community`

```sh
m365 yammer search --queryText "community" --show "groups"
```

Returns topics that match `community`

```sh
m365 yammer search --queryText "community" --show "topics"
```

Returns the first 50 users who match the search query `nuborocks.onmicrosoft.com`

```sh
m365 yammer search --queryText "nuborocks.onmicrosoft.com" --show "users" --limit 50
```

Returns all search results for the query `community`. Stops at 1000 results.

```sh
m365 yammer search --queryText "community" --output json
```

Returns the search results for the query `community` and limits the results to the first 50 entries for each result category.

```sh
m365 yammer search --queryText "community" --output json --limit 50
```

## Response

=== "JSON"

    ```json
    {
      "summary": {
        "messages": 1,
        "topics": 0,
        "users": 0,
        "groups": 0
      },
      "messages": [
        {
          "id": 2000337648877569,
          "sender_id": 36425097217,
          "delegate_id": null,
          "replied_to_id": null,
          "created_at": "2022/11/11 21:00:14 +0000",
          "network_id": 5897756673,
          "message_type": "update",
          "sender_type": "user",
          "url": "https://www.yammer.com/api/v1/messages/2000337648877569",
          "web_url": "https://www.yammer.com/contoso.onmicrosoft.com/messages/2000337648877569",
          "group_id": 31158067201,
          "body": {
            "parsed": "Hello everyone!",
            "plain": "Hello everyone!",
            "rich": "Hello everyone!"
          },
          "thread_id": 2000337648877569,
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
            "count": 1,
            "names": [
              {
                "full_name": "johndoe",
                "permalink": "JohnDoe",
                "user_id": 36425097217,
                "network_id": 5897756673
              }
            ]
          },
          "supplemental_reply": false,
          "content_excerpt": "Hello everyone!",
          "group_created_id": 31158067201
        }
      ],
      "users": [],
      "topics": [],
      "groups": []
    }
    ```

=== "Text"

    ```text
    id                description      type     web_url
    ----------------  ---------------  -------  ----------------------------------------------------------------------------
    2000337648877569  Hello everyone!  message  https://www.yammer.com/contoso.onmicrosoft.com/messages/2000337648877569
    ```

=== "CSV"

    ```csv
    id,description,type,web_url
    2000337648877569,Hello everyone!,message,https://www.yammer.com/contoso.onmicrosoft.com/messages/2000337648877569
    ```
