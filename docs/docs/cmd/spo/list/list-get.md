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
: ID of the list to retrieve information for. Specify either `id`, `title` or `url` but not multiple.

`-t, --title [title]`
: Title of the list to retrieve information for. Specify either `id`, `title` or `url` but not multiple.

`--url [url]`
: Server- or site-relative URL of the list. Specify either `id`, `title` or `url` but not multiple.

`-p, --properties [properties]`
: Comma-separated list of properties to retrieve from the list. Will retrieve all properties possible from default response, if not specified.

`--withPermissions`
: Set if you want to return associated roles and permissions of the list.

--8<-- "docs/cmd/_global.md"

## Examples

Get information about a list with specified ID located in the specified site.

```sh
m365 spo list get --id 0cd891ef-afce-4e55-b836-fce03286cccf --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list with specified title located in the specified site.

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list with specified server relative url located in the specified site.

```sh
m365 spo list get --url 'sites/project-x/Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list with specified site-relative URL located in the specified site.

```sh
m365 spo list get --url 'Shared Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

Get information about a list returning the specified list properties.

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --properties "Title,Id,HasUniqueRoleAssignments,AllowContentTypes"
```

Get information about a list along with the roles and permissions.

```sh
m365 spo list get --title Documents --webUrl https://contoso.sharepoint.com/sites/project-x --withPermissions
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)

