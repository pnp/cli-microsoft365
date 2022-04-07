# spo list label get

Gets label set on the specified list

## Usage

```sh
m365 spo list label get  [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to get the label from is located

`-l, --listId [listId]`
: ID of the list to get the label from. Specify either `listId` or `listTitle` but not both

`-t, --listTitle [listTitle]`
: Title of the list to get the label from. Specify either `listId` or `listTitle` but not both

--8<-- "docs/cmd/_global.md"

## Examples

Gets label set on the list with title _ContosoList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label get  --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle ContosoList
```

Gets label set on the list with id _cc27a922-8224-4296-90a5-ebbc54da2e85_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list label get  --webUrl https://contoso.sharepoint.com/sites/project-x --listId cc27a922-8224-4296-90a5-ebbc54da2e85
```
