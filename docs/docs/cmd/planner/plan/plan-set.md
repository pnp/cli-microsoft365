# planner plan set

Updates a Microsoft Planner plan

## Usage

```sh
m365 planner plan set [options]
```

## Options

`-i, --id [id]`
: ID of the plan. Specify either `id` or `title` but not both.

`-t, --title [title]`
: Title of the plan. Specify either `id` or `title` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

`--ownerGroupName [ownerGroupName]`
: Name of the Group to which the plan belongs. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

`--newTitle [newTitle]`
: New title of the plan.

`--shareWithUserIds [shareWithUserIds]`
: The comma-separated IDs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

`--shareWithUserNames [shareWithUserNames]`
: The comma-separated UPNs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

This command allows using unknown options. 

`--category1 [category1]`
: New label for a category. Define the category key within your option to update the related label. Category 1 to 25 are available. E.g., `--category4`, `--category12`.

## Examples

Updates a Microsoft Planner plan title to New Title

```sh
m365 planner plan set --id 'gndWOTSK60GfPQfiDDj43JgACDCb' --newTitle 'New Title'
```

Share a Microsoft Planner plan owned by the group, with the users

```sh
m365 planner plan set --title 'Plan Title' --ownerGroupName 'My Group' --shareWithUserNames 'user1@contoso.com,user2@contoso.com'
```

Updates a Microsoft Planner plan category labels

```sh
m365 planner plan set --id 'gndWOTSK60GfPQfiDDj43JgACDCb' --category21 'ToDo' --category25 'Urgent'
```

## More information

- Update plannerPlan: [https://learn.microsoft.com/en-us/graph/api/plannerplan-update?view=graph-rest-1.0&tabs=http](https://learn.microsoft.com/en-us/graph/api/plannerplan-update?view=graph-rest-1.0&tabs=http)
- plannerPlanDetails resource type: [https://learn.microsoft.com/en-us/graph/api/resources/plannerplandetails?view=graph-rest-1.0](https://learn.microsoft.com/en-us/graph/api/resources/plannerplandetails?view=graph-rest-1.0)
- plannerCategoryDescriptions resource type: [https://learn.microsoft.com/en-us/graph/api/resources/plannercategorydescriptions?view=graph-rest-1.0](https://learn.microsoft.com/en-us/graph/api/resources/plannercategorydescriptions?view=graph-rest-1.0)


## Response

=== "JSON"

    ```json
    {
      "createdDateTime": "2015-03-30T18:36:49.2407981Z",
      "owner": "ebf3b108-5234-4e22-b93d-656d7dae5874",
      "title": "My Planner Plan",
      "id": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
      "createdBy": {
        "user": {
          "displayName": null,
          "id": "95e27074-6c4a-447a-aa24-9d718a0b86fa"
        },
        "application": {
          "displayName": null,
          "id": "ebf3b108-5234-4e22-b93d-656d7dae5874"
        }
      },
      "container": {
        "containerId": "ebf3b108-5234-4e22-b93d-656d7dae5874",
        "type": "group",
        "url": "https://graph.microsoft.com/v1.0/groups/ebf3b108-5234-4e22-b93d-656d7dae5874"
      },
      "sharedWith": {
        "ebf3b108-5234-4e22-b93d-656d7dae5874": true,
        "6463a5ce-2119-4198-9f2a-628761df4a62": true
      },
      "categoryDescriptions": {
        "category1": null,
        "category2": null,
        "category3": null,
        "category4": null,
        "category5": null,
        "category6": null,
        "category7": null,
        "category8": null,
        "category9": null,
        "category10": null,
        "category11": null,
        "category12": null,
        "category13": null,
        "category14": null,
        "category15": null,
        "category16": null,
        "category17": null,
        "category18": null,
        "category19": null,
        "category20": null,
        "category21": null,
        "category22": null,
        "category23": null,
        "category24": null,
        "category25": null
      }
    }
    ```

=== "Text"

    ```text
    createdDateTime: 2015-03-30T18:36:49.2407981Z
    id             : xqQg5FS2LkCp935s-FIFm2QAFkHM
    owner          : ebf3b108-5234-4e22-b93d-656d7dae5874
    title          : My Planner Plan
    ```

=== "CSV"

    ```csv
    id,title,createdDateTime,owner
    xqQg5FS2LkCp935s-FIFm2QAFkHM,My Planner Plan,2015-03-30T18:36:49.2407981Z,ebf3b108-5234-4e22-b93d-656d7dae5874
    ```

=== "Markdown"

    ```md
    # planner plan set --id "xqQg5FS2LkCp935s-FIFm2QAFkHM" --newTitle "My Planner Plan"

    Date: 27/12/2022

    ## My Planner Plan (xqQg5FS2LkCp935s-FIFm2QAFkHM)

    Property | Value
    ---------|-------
    createdDateTime | 2015-03-30T18:36:49.2407981Z
    owner | ebf3b108-5234-4e22-b93d-656d7dae5874
    title | My Planner Plan
    id | xqQg5FS2LkCp935s-FIFm2QAFkHM
    createdBy | {"user":{"displayName":null,"id":"dd8b99a7-77c6-4238-a609-396d27844921"},"application":{"displayName":null,"id":"31359c7f-bd7e-475c-86db-fdb8c937548e"}}
    container | {"containerId":"ebf3b108-5234-4e22-b93d-656d7dae5874","type":"group","url":"https://graph.microsoft.com/v1.0/groups/ebf3b108-5234-4e22-b93d-656d7dae5874"}
    sharedWith | {"ebf3b108-5234-4e22-b93d-656d7dae5874":true,"6463a5ce-2119-4198-9f2a-628761df4a62":true}
    categoryDescriptions | {"category1":null,"category2":null,"category3":null,"category4":null,"category5":null,"category6":null,"category7":null,"category8":null,"category9":null,"category10":null,"category11":null,"category12":null,"category13":null,"category14":null,"category15":null,"category16":null,"category17":null,"category18":null,"category19":null,"category20":null,"category21":null,"category22":null,"category23":null,"category24":null,"category25":null}
    ```
