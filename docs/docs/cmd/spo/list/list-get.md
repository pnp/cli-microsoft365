# spo list get

Gets information about the specific list

## Usage

```sh
m365 spo list get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to retrieve is located

`-i, --id [id]`
: ID of the list to retrieve information for. Specify either `id` or `title` but not both

`-t, --title [title]`
: Title of the list to retrieve information for. Specify either `id` or `title` but not both

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve from the list. Will retrieve all properties possible from default response, if not specified.

`--withPermissions`
: Set if you want to return associated roles and permissions of the list.

--8<-- "docs/cmd/_global.md"

## Examples

Return information about a list with ID _0cd891ef-afce-4e55-b836-fce03286cccf_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --id 0cd891ef-afce-4e55-b836-fce03286cccf --webUrl https://contoso.sharepoint.com/sites/project-x
```

Return information about a list with title _Documents_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list returning the specified list properties

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Id,HasUniqueRoleAssignments,AllowContentTypes"
```

Get information about a list along with the roles and permissions

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --withPermissions
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)