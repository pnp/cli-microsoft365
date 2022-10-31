# teams user app list

List the apps installed in the personal scope of the specified user

## Usage

```sh
m365 teams user app list [options]
```

## Options

`--userId [userId]`
: The ID of the user to get the apps from. Specify `userId` or `userName` but not both.

`--userName [userName]`
: The UPN of the user to get the apps from. Specify `userId` or `userName` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

List the apps installed in the personal scope of the specified user using its ID

```sh
m365 teams user app list --userId 4440558e-8c73-4597-abc7-3644a64c4bce
```

List the apps installed in the personal scope of the specified user using its UPN

```sh
m365 teams user app list --userName admin@contoso.com
```

## Response

=== "JSON"

    ``` json
    [
      {
        "id": "NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=",
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
        },
        "appId": "14d6962d-6eeb-4f48-8890-de55454bb136"
      },
      {
        "id": "NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMyMGMzNDQwZC1jNjdlLTQ0MjAtOWY4MC0wZTUwYzM5NjkzZGY=",
        "teamsAppDefinition": {
          "id": "MjBjMzQ0MGQtYzY3ZS00NDIwLTlmODAtMGU1MGMzOTY5M2RmIyMxLjAjI1B1Ymxpc2hlZA==",
          "teamsAppId": "20c3440d-c67e-4420-9f80-0e50c39693df",
          "displayName": "Calling",
          "version": "1.0",
          "publishingState": "published",
          "shortDescription": "Calling app bar entry.",
          "description": "Calling app bar entry.",
          "lastModifiedDateTime": null,
          "createdBy": null
        },
        "appId": "20c3440d-c67e-4420-9f80-0e50c39693df"
      }
    ]
    ```

=== "Text"

    ``` text
    id                                                                                                        appId                                    displayName                 version
    --------------------------------------------------------------------------------------------------------  ---------------------------------------  --------------------------  -------
    NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=      14d6962d-6eeb-4f48-8890-de55454bb136     Activity                    1.0
    NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMyMGMzNDQwZC1jNjdlLTQ0MjAtOWY4MC0wZTUwYzM5NjkzZGY=      20c3440d-c67e-4420-9f80-0e50c39693df     Calling                     1.0
    ```

=== "CSV"

    ``` text
    id,appId,displayName,version
    NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMxNGQ2OTYyZC02ZWViLTRmNDgtODg5MC1kZTU1NDU0YmIxMzY=,14d6962d-6eeb-4f48-8890-de55454bb136,Activity,1.0
    NzhjY2Y1MzAtYmJmMC00N2U0LWFhZTYtZGE1ZjhjNmZiMTQyIyMyMGMzNDQwZC1jNjdlLTQ0MjAtOWY4MC0wZTUwYzM5NjkzZGY=,20c3440d-c67e-4420-9f80-0e50c39693df,Calling,1.0
    ```

