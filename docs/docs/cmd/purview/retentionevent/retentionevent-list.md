# purview retentionevent list

Get a list of retention events

## Usage

```sh
m365 purview retentionevent list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get a list of retention events

```sh
m365 purview retentionevent list
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

!!! attention
    This command currently only supports delegated permissions.

## More information

This command is part of a series of commands that have to do with event-based retention. Event-based retention is about starting a retention period when a specific event occurs, instead of the moment a document was labeled or created.

## Response


=== "JSON"

    ```json
    [
      {
        "displayName": "Contract xyz expired",
        "description": null,
        "eventTriggerDateTime": "2023-02-03T13:51:40Z",
        "eventStatus": null,
        "lastStatusUpdateDateTime": null,
        "createdDateTime": "2023-02-03T13:51:40Z",
        "lastModifiedDateTime": "2023-02-03T13:51:40Z",
        "id": "7248cfa8-c03a-4ec1-49a4-08db05edc686",
        "eventQueries": [],
        "eventPropagationResults": [],
        "createdBy": {
          "user": {
            "id": null,
            "displayName": "John Doe"
          }
        },
        "lastModifiedBy": {
          "user": {
            "id": null,
            "displayName": "John Doe"
          }
        }
      }
    ]
    ```

=== "Text"

    ```text
    id                                    displayName            eventTriggerDateTime
    ------------------------------------  ---------------------  --------------------
    7248cfa8-c03a-4ec1-49a4-08db05edc686  Contract xyz expired  2023-02-03T13:51:40Z
    ```

=== "CSV"

    ```csv
    id,displayName,isInUse
    7248cfa8-c03a-4ec1-49a4-08db05edc686,Contract xyz expired,2023-02-03T13:51:40Z
    ```

=== "Markdown"

    ```md
    # purview retentionevent list

    Date: 3/2/2023

    ## Contract xyz expired (7248cfa8-c03a-4ec1-49a4-08db05edc686)

    Property | Value
    ---------|-------
    displayName | Contract xyz expired
    description | null
    eventTriggerDateTime | 2023-02-03T13:51:40Z
    eventStatus | null
    lastStatusUpdateDateTime | null
    createdDateTime | 2023-02-03T13:51:40Z
    lastModifiedDateTime | 2023-02-03T13:51:40Z
    id | 7248cfa8-c03a-4ec1-49a4-08db05edc686
    eventQueries | []
    eventPropagationResults | []
    createdBy | {"user":{"id":null,"displayName":"John Doe"}}
    lastModifiedBy | {"user":{"id":null,"displayName":"John Doe"}}
    ```
