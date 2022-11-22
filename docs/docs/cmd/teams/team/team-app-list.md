# teams team app list

List apps installed in the specified team

## Usage

```sh
m365 teams team app list [options]
```

## Options

`-i, --id [id]`
: The id of the Microsoft Team to list installed apps from. Specify either `id` or `name` but not both.

`-n, --name [name]`
: The name of the Microsoft Team to list installed apps from. Specify either `id` or `name` but not both.


--8<-- "docs/cmd/_global.md"

## Examples

List applications installed in the specified Microsoft Team by id

```sh
m365 teams team app list --id 2eaf7dcd-7e83-4c3a-94f7-932a1299c844
```

List applications installed in the specified Microsoft Team by name

```sh
m365 teams team app list --name "Team Name"
```

## Response

=== "JSON"

    ``` json
    [
       {
        "id": "MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=",
        "teamsApp": {
          "id": "14d6962d-6eeb-4f48-8890-de55454bb136",
          "externalId": null,
          "displayName": "Activity",
          "distributionMethod": "store"
        },
        "teamsAppDefinition": {
          "id": "MTRkNjk2MmQtNmVlYi00ZjQ4LTg4OTAtZGU1NTQ1NGJiMTM2IyMxLjAjI1B1Ymxpc2hlZA==",
          "teamsAppId": "14d6962d-6eeb-4f48-8890-de55454bb136",
          "displayName": "Activity",
          "version": "1.0",
          "publishingState": "published",
          "shortDescription": "Activity app bar entry.",
          "description": "Activity app bar entry.",
          "lastModifiedDateTime": null,
          "createdBy": null
        }
      }
    ]
    ```

=== "Text"

    ``` text
    id                                                                                                    displayName  distributionMethod
    ----------------------------------------------------------------------------------------------------  -----------  ------------------
    MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=  Activity     store
    ```

=== "CSV"

    ``` text
    id,displayName,distributionMethod
    MGFkNTViNWQtNmE3OS00NjdiLWFkMjEtZDRiZWY3OTQ4YTc5IyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=,Activity,store
    ```
