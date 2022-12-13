# pp chatbot list

Lists Microsoft Power Platform chatbot in the specified Power Platform environment

## Usage

```sh
pp chatbot list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List chatbots in a specific environment.

```sh
m365 pp chatbot list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
```

List chatbots in a specific environment as admin.

```sh
m365 pp chatbot list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "language": 1033,
        "botid": "23f5f586-97fd-43d5-95eb-451c9797a53d",
        "authenticationTrigger": 0,
        "stateCode": 0,
        "createdOn": "2022-11-19T10:42:22Z",
        "cdsBotId": "23f5f586-97fd-43d5-95eb-451c9797a53d",
        "schemaName": "new_bot_23f5f58697fd43d595eb451c9797a53d",
        "ownerId": "5fa787c1-1c4d-ed11-bba1-000d3a2caf7f",
        "botModifiedOn": "2022-11-19T20:19:57Z",
        "solutionId": "fd140aae-4df4-11dd-bd17-0019b9312238",
        "isManaged": false,
        "versionNumber": 1429641,
        "timezoneRuleVersionNumber": 0,
        "displayName": "CLI Chatbot",
        "statusCode": 1,
        "owner": "Doe, John",
        "overwriteTime": "1900-01-01T00:00:00Z",
        "componentState": 0,
        "componentIdUnique": "cdcd6496-e25d-4ad1-91cf-3f4d547fdd23",
        "authenticationMode": 1,
        "botModifiedBy": "Doe, John",
        "accessControlPolicy": 0,
        "publishedOn": "2022-11-19T19:19:53Z"
      }
    ]
    ```

=== "Text"

    ```text
    displayName   botid                                 publishedOn           createdOn             botModifiedOn
    ------------  ------------------------------------  --------------------  --------------------  --------------------
    CLI Chatbot   23f5f586-97fd-43d5-95eb-451c9797a53d  2022-11-19T19:19:53Z  2022-11-19T10:42:22Z  2022-11-19T20:19:57Z
    ```

=== "CSV"

    ```csv
    displayName,botid,publishedOn,createdOn,botModifiedOn
    CLI Chatbot,23f5f586-97fd-43d5-95eb-451c9797a53d,2022-11-19T19:19:53Z,2022-11-19T10:42:22Z,2022-11-19T20:19:57Z
    ```
