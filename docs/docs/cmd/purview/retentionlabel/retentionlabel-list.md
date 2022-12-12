# purview retentionlabel list

Get a list of retention labels

## Usage

```sh
m365 purview retentionlabel list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get a list of retention labels

```sh
m365 purview retentionlabel list
```

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

## Response


=== "JSON"

    ```json
    [
      {
        "displayName": "Some label",
        "descriptionForAdmins": "",
        "descriptionForUsers": null,
        "isInUse": true,
        "retentionTrigger": "dateCreated",
        "behaviorDuringRetentionPeriod": "retainAsRecord",
        "actionAfterRetentionPeriod": "delete",
        "createdDateTime": "2022-11-03T10:28:15Z",
        "lastModifiedDateTime": "2022-11-03T10:28:15Z",
        "labelToBeApplied": null,
        "defaultRecordBehavior": "startLocked",
        "id": "dc67203a-6cca-4066-b501-903401308f98",
        "retentionDuration": {
          "days": 365
        },
        "createdBy": {
          "user": {
            "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
            "displayName": null
          }
        },
        "lastModifiedBy": {
          "user": {
            "id": "b52ffd35-d6fe-4b70-86d8-91cc01d76333",
            "displayName": null
          }
        },
        "dispositionReviewStages": []
      }
    ]
    ```

=== "Text"

    ```text
    id                                    displayName     isInUse
    ------------------------------------  --------------  --------
    dc67203a-6cca-4066-b501-903401308f98  Some label      true
    ```

=== "CSV"

    ```csv
    id,displayName,isInUse
    dc67203a-6cca-4066-b501-903401308f98,Some label,true
    ```
