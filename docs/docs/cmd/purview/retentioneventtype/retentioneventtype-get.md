# purview retentioneventtype get

Get a retention event type

## Usage

```sh
m365 purview retentioneventtype get [options]
```

## Options

`-i, --id <id>`
: The Id of the retention event type.

--8<-- "docs/cmd/_global.md"

## Examples

Get a retention event type by id

```sh
m365 purview retentioneventtype get --id c37d695e-d581-4ae9-82a0-9364eba4291e
```

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

!!! attention
    This command currently only supports delegated permissions.

## More information

This command is part of a series of commands that have to do with event-based retention. Event-based retention is about starting a retention period when a specific event occurs, instead of the moment a document was labeled or created.

[Read more on event-based retention here](https://learn.microsoft.com/en-us/microsoft-365/compliance/event-driven-retention?view=o365-worldwide)

## Response

=== "JSON"

    ```json
    {
      "displayName": "Test retention event type",
      "description": "Description for the retention event type",
      "createdDateTime": "2023-01-29T09:30:42Z",
      "lastModifiedDateTime": "2023-01-29T09:30:42Z",
      "id": "c37d695e-d581-4ae9-82a0-9364eba4291e",
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
    ```

=== "Text"

    ```text
    createdBy           : {"user":{"id":null,"displayName":"John Doe"}}
    createdDateTime     : 2023-01-29T09:30:42Z
    description         : Description for the retention event type
    displayName         : Test retention event type
    id                  : c37d695e-d581-4ae9-82a0-9364eba4291e
    lastModifiedBy      : {"user":{"id":null,"displayName":"John Doe"}}
    lastModifiedDateTime: 2023-01-29T09:30:42Z
    ```

=== "CSV"

    ```csv
    displayName,description,createdDateTime,lastModifiedDateTime,id,createdBy,lastModifiedBy
    Test retention event type,Description for the retention event type,2023-01-29T09:30:42Z,2023-01-29T09:30:42Z,c37d695e-d581-4ae9-82a0-9364eba4291e,"{""user"":{""id"":null,""displayName"":""John Doe""}}","{""user"":{""id"":null,""displayName"":""John Doe""}}"
    ```

=== "Markdown"

    ```md
    # purview retentioneventtype get --id "c37d695e-d581-4ae9-82a0-9364eba4291e"

    Date: 1/29/2023

    ## Test retention event type (c37d695e-d581-4ae9-82a0-9364eba4291e)

    Property | Value
    ---------|-------
    displayName | Test retention event type
    description | Description for the retention event type
    createdDateTime | 2023-01-29T09:30:42Z
    lastModifiedDateTime | 2023-01-29T09:30:42Z
    id | c37d695e-d581-4ae9-82a0-9364eba4291e
    createdBy | {"user":{"id":null,"displayName":"John Doe"}}
    lastModifiedBy | {"user":{"id":null,"displayName":"John Doe"}}
    ```
