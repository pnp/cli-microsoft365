# spo list list

Gets all lists within the specified site

## Usage

```sh
m365 spo list list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the lists to retrieve are located

--8<-- "docs/cmd/_global.md"

## Examples

Return all lists located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list list --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- List REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint](https://msdn.microsoft.com/en-us/library/office/dn531433.aspx#bk_ListEndpoint)
