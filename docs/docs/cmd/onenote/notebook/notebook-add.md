# onenote notebook add

Create a new OneNote notebook.

## Usage

```sh
m365 onenote notebook add [options]
```

## Options

`-n, --name <name>`
: Name of the notebook. Notebook names must be unique. The name cannot contain more than 128 characters or contain the following characters: ?\*/:<>

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

--8<-- "docs/cmd/\_global.md"

## Examples

Create a Microsoft OneNote notebook Private Notebook for the currently logged in user

```sh
m365 onenote notebook add --name "Private Notebook"
```

Create a Microsoft OneNote notebook Private Notebook in group 233e43d0-dc6a-482e-9b4e-0de7a7bce9b4

```sh
m365 onenote notebook add --name "Private Notebook" --groupId 233e43d0-dc6a-482e-9b4e-0de7a7bce9b4
```

Create a Microsoft OneNote notebook Private Notebook in group My Group

```sh
m365 onenote notebook add --name "Private Notebook" --groupName "MyGroup"
```

Create a Microsoft OneNote notebook Private Notebook for user user1@contoso.onmicrosoft.com

```sh
m365 onenote notebook add --name "Private Notebook" --userName user1@contoso.onmicrosoft.com
```

Create a Microsoft OneNote notebook Private Notebook for user 2609af39-7775-4f94-a3dc-0dd67657e900

```sh
m365 onenote notebook add --name "Private Notebook" --userId 2609af39-7775-4f94-a3dc-0dd67657e900
```

Create a Microsoft OneNote notebook Private Notebook for site https://contoso.sharepoint.com/sites/testsite

```sh
m365 onenote notebook add --name "Private Notebook" --webUrl https://contoso.sharepoint.com/sites/testsite
```

## More information

- Create notebook (MS Graph docs): [https://docs.microsoft.com/en-us/graph/api/onenote-post-notebooks?view=graph-rest-1.0&tabs=http](https://docs.microsoft.com/en-us/graph/api/onenote-post-notebooks?view=graph-rest-1.0&tabs=http)
