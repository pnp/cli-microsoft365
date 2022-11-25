# onenote notebook list

Retrieve a list of notebooks.

## Usage

```sh
m365 onenote notebook list [options]
```

## Options

`--userId [userId]`
: Id of the user. Use either userId or userName, but not both.

`--userName [userName]`
: Name of the user. Use either userId or userName, but not both.

`--groupId [groupId]`
: Id of the SharePoint group. Use either groupName or groupId, but not both

`--groupName [groupName]`
: Name of the SharePoint group. Use either groupName or groupId, but not both.

`-u, --webUrl [webUrl]`
: URL of the SharePoint site.

--8<-- "docs/cmd/_global.md"

## Examples

List Microsoft OneNote notebooks for the currently logged in user

```sh
m365 onenote notebook list
```

List Microsoft OneNote notebooks in group 233e43d0-dc6a-482e-9b4e-0de7a7bce9b4

```sh
m365 onenote notebook list --groupId 233e43d0-dc6a-482e-9b4e-0de7a7bce9b4
```

List Microsoft OneNote notebooks in group My Group

```sh
m365 onenote notebook list --groupName "MyGroup"
```

List Microsoft OneNote notebooks for user user1@contoso.onmicrosoft.com

```sh
m365 onenote notebook list --userName user1@contoso.onmicrosoft.com
```

List Microsoft OneNote notebooks for user 2609af39-7775-4f94-a3dc-0dd67657e900

```sh
m365 onenote notebook list --userId 2609af39-7775-4f94-a3dc-0dd67657e900
```

List Microsoft OneNote notebooks for site https://contoso.sharepoint.com/sites/testsite

```sh
m365 onenote notebook list --webUrl https://contoso.sharepoint.com/sites/testsite
```

## More information

- List notebooks (MS Graph docs): [https://docs.microsoft.com/en-us/graph/api/onenote-list-notebooks?view=graph-rest-1.0&tabs=http](https://docs.microsoft.com/en-us/graph/api/onenote-list-notebooks?view=graph-rest-1.0&tabs=http)

## Response

=== "JSON"

    ```json
    [
      {
        "id": "1-db46efb2-1b80-4a28-987d-507cff465418",
        "self": "https://graph.microsoft.com/v1.0/users/cfa02cf8-e903-4b76-a17b-15e43d4cf81e/onenote/notebooks/1-db46efb2-1b80-4a28-987d-507cff465418",
        "createdDateTime": "2017-11-24T11:00:46Z",
        "displayName": "TeamsNotebook(Shared)",
        "lastModifiedDateTime": "2017-11-24T11:00:46Z",
        "isDefault": false,
        "userRole": "Owner",
        "isShared": false,
        "sectionsUrl": "https://graph.microsoft.com/v1.0/users/cfa02cf8-e903-4b76-a17b-15e43d4cf81e/onenote/notebooks/1-db46efb2-1b80-4a28-987d-507cff465418/sections",
        "sectionGroupsUrl": "https://graph.microsoft.com/v1.0/users/cfa02cf8-e903-4b76-a17b-15e43d4cf81e/onenote/notebooks/1-db46efb2-1b80-4a28-987d-507cff465418/sectionGroups",
        "createdBy": {
          "user": {
            "id": "cfa02cf8-e903-4b76-a17b-15e43d4cf81e",
            "displayName": "Demo User1"
          }
        },
        "lastModifiedBy": {
          "user": {
            "id": "cfa02cf8-e903-4b76-a17b-15e43d4cf81e",
            "displayName": "Demo User1"
          }
        },
        "links": {
          "oneNoteClientUrl": {
            "href": "onenote:https://contoso-my.sharepoint.com/personal/user1_contoso_com/Documents/Notebooks/TeamsNotebook(Shared)"
          },
          "oneNoteWebUrl": {
            "href": "https://contoso-my.sharepoint.com/personal/user1_contoso_com/Documents/Notebooks/TeamsNotebook(Shared)"
          }
        }
      }
    ]
    ```

=== "Text"

    ```text
    createdDateTime: 2017-11-24T11:00:46Z
    displayName : TeamsNotebook(Shared)
    id : 1-db46efb2-1b80-4a28-987d-507cff465418
    ```

=== "CSV"

    ```csv
    createdDateTime,displayName,id
    2017-11-24T11:00:46Z,TeamsNotebook(Shared),1-db46efb2-1b80-4a28-987d-507cff465418
    ```
