# planner plan get

Retrieve information about the specified plan

## Usage

```sh
m365 planner plan get [options]
```

## Options

`-i, --id [id]`
: ID of the plan. Specify either `id` or `title` but not both.

`-t, --title [title]`
: Title of the plan. Specify either `id` or `title` but not both.

`--ownerGroupId [ownerGroupId]`
: ID of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

`--ownerGroupName [ownerGroupName]`
: Name of the Group that owns the plan. Specify either `ownerGroupId` or `ownerGroupName` when using `title`.

--8<-- "docs/cmd/_global.md"

## Examples

Returns the Microsoft Planner plan with id _gndWOTSK60GfPQfiDDj43JgACDCb_

```sh
m365 planner plan get --id "gndWOTSK60GfPQfiDDj43JgACDCb"
```

Returns the Microsoft Planner plan with title _MyPlan_ for Group _233e43d0-dc6a-482e-9b4e-0de7a7bce9b4_

```sh
m365 planner plan get --title "MyPlan" --ownerGroupId "233e43d0-dc6a-482e-9b4e-0de7a7bce9b4"
```

Returns the Microsoft Planner plan with title _MyPlan_ for Group _My Planner Group_

```sh
m365 planner plan get --title "MyPlan" --ownerGroupName "My Planner Group"
```

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
    # planner plan get --id "xqQg5FS2LkCp935s-FIFm2QAFkHM"

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
    sharedWith | {"aa089c1a-0ca8-4b36-ae2e-c5d60f054a2c":true,"ebf3b108-5234-4e22-b93d-656d7dae5874":true}
    categoryDescriptions | {"category1":null,"category2":null,"category3":null,"category4":null,"category5":null,"category6":null,"category7":null,"category8":null,"category9":null,"category10":null,"category11":null,"category12":null,"category13":null,"category14":null,"category15":null,"category16":null,"category17":null,"category18":null,"category19":null,"category20":null,"category21":null,"category22":null,"category23":null,"category24":null,"category25":null}
    ```
