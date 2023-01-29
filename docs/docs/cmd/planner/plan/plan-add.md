# planner plan add

Adds a new Microsoft Planner plan

## Usage

```sh
m365 planner plan add [options]
```

## Options

`-t, --title <title>`
: Title of the plan to add.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. A valid group must exist before this option can be set. Specify either `ownerGroupId`, `ownerGroupName` or `rosterId`.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. A valid group must exist before this option can be set. Specify either `ownerGroupId`, `ownerGroupName` or `rosterId`.

`--rosterId [rosterId]`
: ID of the Planner Roster. Specify either `ownerGroupId`, `ownerGroupName` or `rosterId`.

`--shareWithUserIds [shareWithUserIds]`
: The comma-separated IDs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

`--shareWithUserNames [shareWithUserNames]`
: The comma-separated UPNs of the users with whom you want to share the plan. Specify either `shareWithUserIds` or `shareWithUserNames` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

- Related to the options `--shareWithUserIds` and `--shareWithUserNames`. If you are leveraging Microsoft 365 groups, use the `aad o365group user` commands to manage group membership to share the [group's](https://pnp.github.io/cli-microsoft365/cmd/aad/o365group/o365group-user-add/) plan. You can also add existing members of the group to this collection though it is not required for them to access the plan owned by the group.

- Hint: Unlike for groups, a Planner Roster can contain only 1 plan.

## Examples

Adds a Microsoft Planner plan with a Group by id

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupId '233e43d0-dc6a-482e-9b4e-0de7a7bce9b4'
```

Adds a Microsoft Planner plan with with a Group by name

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupName 'My Planner Group'
```

Adds a Microsoft Planner plan with with a Group by name and share it with the given users

```sh
m365 planner plan add --title 'My Planner Plan' --ownerGroupName 'My Planner Group' --shareWithUserNames 'Allan.Carroll@contoso.com,Ida.Stevens@contoso.com'
```

## Response

### Standard response

=== "JSON"

    ``` json
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
      }
    }
    ```

=== "Text"

    ``` text
    createdDateTime: 2015-03-30T18:36:49.2407981Z
    id             : xqQg5FS2LkCp935s-FIFm2QAFkHM
    owner          : ebf3b108-5234-4e22-b93d-656d7dae5874
    title          : My Planner Plan
    ```

=== "CSV"

    ``` text
    id,title,createdDateTime,owner
    xqQg5FS2LkCp935s-FIFm2QAFkHM,My Planner Plan,2015-03-30T18:36:49.2407981Z,ebf3b108-5234-4e22-b93d-656d7dae5874
    ```

=== "Markdown"

    ```md
    # planner plan add --title "My Planner Plan" --ownerGroupId "ebf3b108-5234-4e22-b93d-656d7dae5874"

    Date: 30/03/2015

    ## My Planner Plan (xqQg5FS2LkCp935s-FIFm2QAFkHM)

    Property | Value
    ---------|-------
    createdDateTime | 2015-03-30T18:36:49.2407981Z
    owner | ebf3b108-5234-4e22-b93d-656d7dae5874
    title | My Planner Plan
    id | xqQg5FS2LkCp935s-FIFm2QAFkHM
    createdBy | {"user":{"displayName":null,"id":"95e27074-6c4a-447a-aa24-9d718a0b86fa"},"application":{"displayName":null,"id":"ebf3b108-5234-4e22-b93d-656d7dae5874"}}
    container | {"containerId":"ebf3b108-5234-4e22-b93d-656d7dae5874","type":"group","url":"https://graph.microsoft.com/v1.0/groups/ebf3b108-5234-4e22-b93d-656d7dae5874"}
    ```

### `shareWithUserIds`, `shareWithUserNames` response

When we make use of the option `shareWithUserIds` or `shareWithUserNames` the response will differ. 

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
    # planner plan add --title "My Planner Plan" --ownerGroupId "ebf3b108-5234-4e22-b93d-656d7dae5874" --shareWithUserNames "Allan.Carroll@contoso.com,Ida.Stevens@contoso.com"

    Date: 30/03/2015

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
