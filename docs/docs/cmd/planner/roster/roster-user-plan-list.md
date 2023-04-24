# planner roster user plan list

Lists all Microsoft Planner Roster plans for a specified user.

## Usage

```sh
m365 planner roster user plan list [options]
```

## Options

`--userId [userId]`
: User's Azure AD ID. Specify either `userId` or `userName` but not both. Specify this option only when using application permissions.

`--userName [userName]`
: User's UPN (user principal name, e.g. johndoe@example.com). Specify either `userId` or `userName` but not both. Specify this option only when using application permissions.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

List all Planner plans contained in a Roster where the current logged in user is member of.

```sh
m365 planner roster user plan list
```

List all Planner plans contained in a Roster where specific user is member of by its UPN.

```sh
m365 planner roster user plan list --userName john.doe@contoso.com
```

List all Planner plans contained in a Roster where specific user is member of by its Id.

```sh
m365 planner roster user plan list --userId 59f80e08-24b1-41f8-8586-16765fd830d3
```

## Response

=== "JSON"

    ```json
    [
      {
        "createdDateTime": "2023-04-06T14:41:49.8676617Z",
        "owner": "59f80e08-24b1-41f8-8586-16765fd830d3",
        "title": "My Planner Plan",
        "creationSource": null,
        "id": "_5GY9MJpZU2vb3DC46CP3MkACr8m",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "59f80e08-24b1-41f8-8586-16765fd830d3"
          },
          "application": {
            "displayName": null,
            "id": "31359c7f-bd7e-475c-86db-fdb8c937548e"
          }
        },
        "container": {
          "containerId": "_5GY9MJpZU2vb3DC46CP3MkACr8m",
          "type": "unknownFutureValue",
          "url": "https://graph.microsoft.com/beta/planner/rosters/_5GY9MJpZU2vb3DC46CP3MkACr8m"
        },
        "contexts": {},
        "sharedWithContainers": []
      }
    ]
    ```

=== "Text"

    ```text
    createdDateTime: 2023-04-06T14:41:49.8676617Z
    id             : _5GY9MJpZU2vb3DC46CP3MkACr8m
    owner          : 59f80e08-24b1-41f8-8586-16765fd830d3
    title          : My Planner Plan
    ```

=== "CSV"

    ```csv
    createdDateTime,owner,title,id
    2023-04-06T14:41:49.8676617Z,59f80e08-24b1-41f8-8586-16765fd830d3,My Planner Plan,_5GY9MJpZU2vb3DC46CP3MkACr8m
    ```

=== "Markdown"

    ```md
    # planner roster user plan list --userId "59f80e08-24b1-41f8-8586-16765fd830d3"

    Date: 4/8/2023

    ## My Planner Plan (_5GY9MJpZU2vb3DC46CP3MkACr8m)

    Property | Value
    ---------|-------
    createdDateTime | 2023-04-06T14:41:49.8676617Z
    owner | 59f80e08-24b1-41f8-8586-16765fd830d3
    title | My Planner Plan
    id | \_5GY9MJpZU2vb3DC46CP3MkACr8m
    ```
