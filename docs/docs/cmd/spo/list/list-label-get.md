# spo list label get

Gets label set on the specified list

## Usage

```sh
m365 spo list label get  [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to get the label from is located.

`-l, --listId [listId]`
: ID of the list to get the label from. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list to get the label from. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Gets label set on the list with specified title located in the specified site.

```sh
m365 spo list label get --listTitle ContosoList --webUrl https://contoso.sharepoint.com/sites/project-x
```

Gets label set on the list with specified id located in the specified site.

```sh
m365 spo list label get --listId cc27a922-8224-4296-90a5-ebbc54da2e85 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Gets label set on the list with specified server relative url located in the specified site.

```sh
m365 spo list label get --listUrl 'sites/project-x/Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

Gets label set on the list with specified site-relative URL located in the specified site.

```sh
m365 spo list label get --listUrl 'Shared Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

