# planner task add

Adds a new Microsoft Planner task

## Usage

```sh
m365 planner task add [options]
```

## Options

`-t, --title <title>`
: Title of the task to add.

`--planId [planId]`
: ID of the plan to which the task belongs. Specify either `planId`, `planTitle`, or `rosterId` but not multiple.

`--planTitle [planTitle]`
: Title of the plan to which the task belongs. Specify either `planId`, `planTitle`, or `rosterId` but not multiple.

`--rosterId [rosterId]`
: ID of the Planner Roster. Specify either `planId`, `planTitle`, or `rosterId` but not multiple.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--ownerGroupName [ownerGroupName]`
: Name of the group to which the plan belongs. Specify `ownerGroupId` or `ownerGroupName` when using `planTitle`.

`--bucketId [bucketId]`
: ID of the bucket to which the task belongs. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--bucketName [bucketName]`
: Name of the bucket to which the task belongs. The bucket needs to exist in the selected plan. Specify either `bucketId` or `bucketName` but not both.

`--startDateTime [startDateTime]`
: The date and time when the task started. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--dueDateTime [dueDateTime]`
: The date and time when the task is due. This should be defined as a valid ISO 8601 string. `2021-12-16T18:28:48.6964197Z`

`--percentComplete [percentComplete]`
: Percentage of task completion. Number between 0 and 100.

`--assignedToUserIds [assignedToUserIds]`
: The comma-separated IDs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assignedToUserNames [assignedToUserNames]`
: The comma-separated UPNs of the assignees the task is assigned to. Specify either `assignedToUserIds` or `assignedToUserNames` but not both.

`--assigneePriority [assigneePriority]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--description [description]`
: Description of the task.

`--appliedCategories [appliedCategories]`
: Comma-separated categories that should be added to the task. The possible options are: `category1`, `category2`, `category3`, `category4`, `category5` and/or `category6`. Additional info defined [here](https://docs.microsoft.com/graph/api/resources/plannerappliedcategories?view=graph-rest-1.0).

`--previewType [previewType]`
: This sets the type of preview that shows up on the task. The possible values are: `automatic`, `noPreview`, `checklist`, `description`, `reference`. When set to automatic the displayed preview is chosen by the app viewing the task. Default `automatic`.

`--orderHint [orderHint]`
: Hint used to order items of this type in a list view. The format is defined as outlined [here](https://docs.microsoft.com/graph/api/resources/planner-order-hint-format?view=graph-rest-1.0).

`--priority [priority]`
: Priority of the task: Urgent, Important, Medium, Low. Or an integer between 0 and 10 (check remarks section for more info). Default value is Medium.

--8<-- "docs/cmd/_global.md"

## Remarks

When you specify the value for `percentComplete`, consider the following:

- when set to 0, the task is considered _Not started_
- when set between 1 and 99, the task is considered _In progress_
- when set to 100, the task is considered _Completed_

When you specify an integer value for `priority`, consider the following:

- values 0 and 1 are interpreted as _Urgent_
- values 2, 3 and 4 are interpreted as _Important_
- values 5, 6 and 7 are interpreted as _Medium_
- values 8, 9 and 10 are interpreted as _Low_

When using `description` with a multiple lines value, use the new line character of the shell you are using to indicate line breaks. For PowerShell this is `` `n ``. For Zsh or Bash use `\n` with a `$` in front. E.g. `$"Line 1\nLine 2"`.

!!! attention
    When using `rosterId`, the command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Adds a Microsoft Planner task with the name for plan with the specified ID and specified bucket with the ID.

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna"
```

Adds a Completed Microsoft Planner task with the name for plan with the specified title owned by specified group and the specified bucket with the ID.

```sh
m365 planner task add --title "My Planner task" --planTitle "My Planner Plan" --ownerGroupName "My Planner Group" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --percentComplete 100
```

Adds a Microsoft Planner task with the specified name for plan with the specified ID and bucket with the ID. The new task will be assigned to the specified users and receive a specified due date.

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --assignedToUserNames "Allan.Carroll@contoso.com,Ida.Stevens@contoso.com" --dueDateTime "2021-12-16"
```

Adds a Microsoft Planner task with the specified name for plan with the specified ID and bucket with the ID. The new task will be assigned to the specified users who will appear first with the asssignee priority _' !'_ 

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --assignedToUserNames "Allan.Carroll@contoso.com,Ida.Stevens@contoso.com" --asssigneePriority ' !'
```

Adds a Microsoft Planner task with the specified name for plan with the specified ID and the bucket with the ID. The new task will receive the specified categories and get a specified preview with the type.

```sh
m365 planner task add --title "My Planner Task" --planId "8QZEH7b3wkSbGQobscsM5gADCBa" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna" --appliedCategories "category1,category3" --previewType "noPreview"
```

Adds a Microsoft Planner task with the specified name for plan with the specified rosterId and bucket with the ID.

```sh
m365 planner task add --title "My Planner Task" --rosterId "DjL5xiKO10qut8LQgztpKskABWna" --bucketId "IK8tuFTwQEa5vTonM7ZMRZgAKdna"
```

## Response

### Standard response

=== "JSON"

    ```json
    {
      "planId": "oUHpnKBFekqfGE_PS6GGUZcAFY7b",
      "bucketId": "vncYUXCRBke28qMLB-d4xJcACtNz",
      "title": "Important task",
      "orderHint": "8585269241124027581",
      "assigneePriority": "",
      "percentComplete": 50,
      "startDateTime": "2023-01-20T00:00:00Z",
      "createdDateTime": "2023-01-25T21:39:33.0748226Z",
      "dueDateTime": "2023-02-15T00:00:00Z",
      "hasDescription": false,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "priority": 5,
      "id": "D-ys8Ef4kEuwYG4r68Um3pcAAe9M",
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
    ```

=== "Text"

    ```text
    activeChecklistItemCount: 0
    appliedCategories       : {}
    assigneePriority        :
    assignments             : {}
    bucketId                : vncYUXCRBke28qMLB-d4xJcACtNz
    checklistItemCount      : 0
    completedBy             : null
    completedDateTime       : null
    conversationThreadId    : null
    createdBy               : {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    createdDateTime         : 2023-01-25T21:44:10.6044385Z
    dueDateTime             : 2023-02-15T00:00:00Z
    hasDescription          : false
    id                      : D-ys8Ef4kEuwYG4r68Um3pcAAe9M
    orderHint               : 8585269238348731422
    percentComplete         : 50
    planId                  : oUHpnKBFekqfGE_PS6GGUZcAFY7b
    previewType             : automatic
    priority                : 5
    referenceCount          : 0
    references              : {}
    startDateTime           : 2023-01-20T00:00:00Z
    title                   : Important task
    ```

=== "CSV"

    ```csv
    planId,bucketId,title,orderHint,assigneePriority,percentComplete,startDateTime,createdDateTime,dueDateTime,hasDescription,previewType,completedDateTime,completedBy,referenceCount,checklistItemCount,activeChecklistItemCount,conversationThreadId,priority,id,createdBy,appliedCategories,assignments
    oUHpnKBFekqfGE_PS6GGUZcAFY7b,vncYUXCRBke28qMLB-d4xJcACtNz,Important task,8585269237867589640,,50,2023-01-20T00:00:00Z,2023-01-25T21:44:58.7186167Z,2023-02-15T00:00:00Z,,automatic,,,0,0,0,,5,D-ys8Ef4kEuwYG4r68Um3pcAAe9M,"{""user"":{""displayName"":null,""id"":""b2091e18-7882-4efe-b7d1-90703f5a5c65""},""application"":{""displayName"":null,""id"":""31359c7f-bd7e-475c-86db-fdb8c937548e""}}",{},{}
    ```

=== "Markdown"

    ```md
    # planner task add --planId "oUHpnKBFekqfGE_PS6GGUZcAFY7b" --bucketName "To do" --startDateTime "2023-01-20" --dueDateTime "2023-02-15" --percentComplete "50" --title "Important task"

    Date: 25/1/2023

    ## Important task (D-ys8Ef4kEuwYG4r68Um3pcAAe9M)

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
    hasDescription | false
    previewType | automatic
    completedDateTime | null
    completedBy | null
    referenceCount | 0
    checklistItemCount | 0
    activeChecklistItemCount | 0
    conversationThreadId | null
    priority | 5
    id | D-ys8Ef4kEuwYG4r68Um3pcAAe9M
    createdBy | {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    appliedCategories | {}
    assignments | {}
    ```

### `description`, `previewType` response

=== "JSON"

    ```json
    {
      "planId": "oUHpnKBFekqfGE_PS6GGUZcAFY7b",
      "bucketId": "vncYUXCRBke28qMLB-d4xJcACtNz",
      "title": "Important task",
      "orderHint": "8585269241124027581",
      "assigneePriority": "",
      "percentComplete": 50,
      "startDateTime": "2023-01-20T00:00:00Z",
      "createdDateTime": "2023-01-25T21:39:33.0748226Z",
      "dueDateTime": "2023-02-15T00:00:00Z",
      "hasDescription": true,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "priority": 5,
      "id": "D-ys8Ef4kEuwYG4r68Um3pcAAe9M",
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
      "assignments": {},
      "description": "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
      "references": {},
      "checklist": {}
    }
    ```

=== "Text"

    ```txt
    activeChecklistItemCount: 0
    appliedCategories       : {}
    assigneePriority        :
    assignments             : {}
    bucketId                : vncYUXCRBke28qMLB-d4xJcACtNz
    checklist               : {}
    checklistItemCount      : 0
    completedBy             : null
    completedDateTime       : null
    conversationThreadId    : null
    createdBy               : {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    createdDateTime         : 2023-01-25T21:44:10.6044385Z
    description             : Lorem ipsum dolor sit amet, consectetur adipiscing elit.
    dueDateTime             : 2023-02-15T00:00:00Z
    hasDescription          : true
    id                      : D-ys8Ef4kEuwYG4r68Um3pcAAe9M
    orderHint               : 8585269238348731422
    percentComplete         : 50
    planId                  : oUHpnKBFekqfGE_PS6GGUZcAFY7b
    previewType             : automatic
    priority                : 5
    referenceCount          : 0
    references              : {}
    startDateTime           : 2023-01-20T00:00:00Z
    title                   : Important task
    ```

=== "CSV"

    ```csv
    planId,bucketId,title,orderHint,assigneePriority,percentComplete,startDateTime,createdDateTime,dueDateTime,hasDescription,previewType,completedDateTime,completedBy,referenceCount,checklistItemCount,activeChecklistItemCount,conversationThreadId,priority,id,createdBy,appliedCategories,assignments,description,references,checklist
    oUHpnKBFekqfGE_PS6GGUZcAFY7b,vncYUXCRBke28qMLB-d4xJcACtNz,Important task,8585269237867589640,,50,2023-01-20T00:00:00Z,2023-01-25T21:44:58.7186167Z,2023-02-15T00:00:00Z,1,automatic,,,0,0,0,,5,D-ys8Ef4kEuwYG4r68Um3pcAAe9M,"{""user"":{""displayName"":null,""id"":""b2091e18-7882-4efe-b7d1-90703f5a5c65""},""application"":{""displayName"":null,""id"":""31359c7f-bd7e-475c-86db-fdb8c937548e""}}",{},{},"Lorem ipsum dolor sit amet, consectetur adipiscing elit.",{},{}
    ```

=== "Markdown"

    ```md
    # planner task add --planId "oUHpnKBFekqfGE_PS6GGUZcAFY7b" --bucketName "To do" --startDateTime "2023-01-20" --dueDateTime "2023-02-15" --percentComplete "50" --title "Important task" --description "Lorem ipsum dolor sit amet, consectetur adipiscing elit."

    Date: 25/1/2023

    ## Important task (D-ys8Ef4kEuwYG4r68Um3pcAAe9M)

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
    checklistItemCount | 0
    activeChecklistItemCount | 0
    conversationThreadId | null
    priority | 5
    id | D-ys8Ef4kEuwYG4r68Um3pcAAe9M
    createdBy | {"user":{"displayName":null,"id":"b2091e18-7882-4efe-b7d1-90703f5a5c65"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    appliedCategories | {}
    assignments | {}
    description | Lorem ipsum dolor sit amet, consectetur adipiscing elit.
    references | {}
    checklist | {}
    ```
