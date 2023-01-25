# planner task list

Lists planner tasks in a bucket, plan, or tasks for the currently logged in user

## Usage

```sh
m365 planner task list [options]
```

## Options

`--bucketId [bucketId]`
: ID of the bucket to list the tasks of. To retrieve tasks from a bucket, specify `bucketId` or `bucketName`, but not both.

`--bucketName [bucketName]`
: Name of the bucket to list the tasks of. To retrieve tasks from a bucket, specify `bucketId` or `bucketName`, but not both.

`--planId [planId]`
: ID of a plan to list the tasks of. To retrieve all tasks from a plan, specify either `planId` or `planTitle` but not both. Use in combination with `bucketName` to retrieve tasks from a specific bucket.

`--planTitle [planTitle]`
: Title of a plan to list the tasks of. To retrieve all tasks from a plan, specify either `planId` or `planTitle` but not both. Always use in combination with either `ownerGroupId` or `ownerGroupName`. Use in combination with `bucketName` to retrieve tasks from a specific bucket.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command uses API that is currently in preview to enrich the results with the `priority` field. Keep in mind that this preview API is subject to change once the API reached general availability.

## Examples

List tasks for the currently logged in user

```sh
m365 planner task list
```

List the Microsoft Planner tasks in the plan _iVPMIgdku0uFlou-KLNg6MkAE1O2_

```sh
m365 planner task list --planId "iVPMIgdku0uFlou-KLNg6MkAE1O2"`
```

List the Microsoft Planner tasks in the plan _My Plan_ in group _My Group_

```sh
m365 planner task list --planTitle "My Plan" --ownerGroupName "My Group"
```

List the Microsoft Planner tasks in the bucket _FtzysDykv0-9s9toWiZhdskAD67z_

```sh
m365 planner task list --bucketId "FtzysDykv0-9s9toWiZhdskAD67z"
```

List the Microsoft Planner tasks in the bucket _My Bucket_ belonging to plan _iVPMIgdku0uFlou-KLNg6MkAE1O2_

```sh
m365 planner task list --bucketName "My Bucket" --planId "iVPMIgdku0uFlou-KLNg6MkAE1O2"
```

List the Microsoft Planner tasks in the bucket _My Bucket_ belonging to plan _My Plan_ in group _My Group_

```sh
m365 planner task list --bucketName "My Bucket" --planTitle "My Plan" --ownerGroupName "My Group"
```

## Response

=== "JSON"

    ```json
    [
      {
        "planId": "oUHpnKBFekqfGE_PS6GGUZcAFY7b",
        "bucketId": "vncYUXCRBke28qMLB-d4xJcACtNz",
        "title": "Important task",
        "orderHint": "8585269235419217847",
        "assigneePriority": "",
        "percentComplete": 50,
        "startDateTime": "2023-01-20T00:00:00Z",
        "createdDateTime": "2023-01-25T21:49:03.555796Z",
        "dueDateTime": "2023-02-15T00:00:00Z",
        "hasDescription": true,
        "previewType": "automatic",
        "completedDateTime": null,
        "completedBy": null,
        "referenceCount": 0,
        "checklistItemCount": 1,
        "activeChecklistItemCount": 1,
        "conversationThreadId": null,
        "priority": 5,
        "id": "OopX1ANphEu7Lm4-0tVtl5cAFRGQ",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "b2091e18-7882-4efe-b7d1-90703f5a5c65"
          },
          "application": {
            "displayName": null,
            "id": "31359c7f-bd7e-475c-86db-fdb8c937548e"
          }
        },
        "appliedCategories": {},
        "assignments": {}
      }
    ]
    ```

=== "Text"

    ```txt
    id                            title           startDateTime         dueDateTime           completedDateTime
    ----------------------------  --------------  --------------------  --------------------  -----------------
    OopX1ANphEu7Lm4-0tVtl5cAFRGQ  Important task  2023-01-20T00:00:00Z  2023-02-15T00:00:00Z  null
    ```

=== "CSV"

    ```csv
    id,title,startDateTime,dueDateTime,completedDateTime
    OopX1ANphEu7Lm4-0tVtl5cAFRGQ,Important task,2023-01-20T00:00:00Z,2023-02-15T00:00:00Z,
    ```

=== "Markdown"

    ```md
    # planner task list --planId "oUHpnKBFekqfGE_PS6GGUZcAFY7b" --bucketName "To do"

    Date: 25/1/2023

    ## Important task (OopX1ANphEu7Lm4-0tVtl5cAFRGQ)

    Property | Value
    ---------|-------
    planId | oUHpnKBFekqfGE\_PS6GGUZcAFY7b
    bucketId | vncYUXCRBke28qMLB-d4xJcACtNz
    title | Important task
    orderHint | 8585269235419217847
    assigneePriority |
    percentComplete | 50
    startDateTime | 2023-01-20T00:00:00Z
    createdDateTime | 2023-01-25T21:49:03.555796Z
    dueDateTime | 2023-02-15T00:00:00Z
    hasDescription | true
    previewType | automatic
    completedDateTime | null
    completedBy | null
    referenceCount | 0
    checklistItemCount | 1
    activeChecklistItemCount | 1
    conversationThreadId | null
    priority | 5
    id | OopX1ANphEu7Lm4-0tVtl5cAFRGQ
    createdBy | {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    appliedCategories | {}
    assignments | {}
    ```
