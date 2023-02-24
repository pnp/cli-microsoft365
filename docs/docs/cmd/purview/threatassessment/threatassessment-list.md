# purview threatassessment list

Get a list of threat assessments

## Usage

```sh
m365 purview threatassessment list [options]
```

## Options

`-t, --type [type]`
: The type of threat assessment to retrieve. Supports `mail`, `file`, `emailFile` and `url`.

--8<-- "docs/cmd/_global.md"

## Examples

Get a list of threat assessments

```sh
m365 purview threatassessment list
```

Get a list of threat assessments of type _mail_

```sh
m365 purview threatassessment list --type mail
```

## Response


=== "JSON"

    ```json
    [
      {
        "@odata.type": "#microsoft.graph.mailAssessmentRequest",
        "id": "49c5ef5b-1f65-444a-e6b9-08d772ea2059",
        "createdDateTime": "2019-11-27T03:30:18.6890937Z",
        "contentType": "mail",
        "expectedAssessment": "block",
        "category": "spam",
        "status": "pending",
        "requestSource": "administrator",
        "recipientEmail": "john@contoso.onmicrosoft.com",
        "destinationRoutingReason": "notJunk",
        "messageUri": "https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=",
        "createdBy": {
          "user": {
            "id": "c52ce8db-3e4b-4181-93c4-7d6b6bffaf60",
            "displayName": "John Doe"
          }
        }
      }
    ];
    ```

=== "Text"

    ```text
    id                                    contentType  category
    ------------------------------------  -----------  --------
    49c5ef5b-1f65-444a-e6b9-08d772ea2059  mail         spam
    ```

=== "CSV"

    ```csv
    id,contentType,category
    49c5ef5b-1f65-444a-e6b9-08d772ea2059,mail,spam
    ```

=== "Markdown"

    ```md
    # purview threatassessment list

    Date: 16/2/2023

    ## a47e428c-a7bd-4cf2-f061-08db0f58b736

    Property | Value
    ---------|-------
    @odata.type | #microsoft.graph.mailAssessmentRequest
    id | 49c5ef5b-1f65-444a-e6b9-08d772ea2059
    createdDateTime | 2019-11-27T03:30:18.6890937Z
    contentType | mail
    expectedAssessment | block
    category | spam
    status | pending
    recipientEmail | john@contoso.onmicrosoft.com
    destinationRoutingReason | notJunk
    messageUri | https://graph.microsoft.com/v1.0/users/c52ce8db-3e4b-4181-93c4-7d6b6bffaf60/messages/AAMkADU3MWUxOTU0LWNlOTEt=
    createdBy | {"user":{"id":"c52ce8db-3e4b-4181-93c4-7d6b6bffaf60","displayName":"John Doe"}}
    ```
