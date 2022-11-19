# spo list sitescript get

Extracts a site script from a SharePoint list

## Usage

```sh
m365 spo list sitescript get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to extract the site script from is located.

`-l, --listId [listId]`
: ID of the list to extract the site script from. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list to extract the site script from. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId`, `listTitle` or `listUrl` but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Extract a site script from an existing SharePoint list with specified title located in the specified site.

```sh
m365 spo list sitescript get --listTitle ContosoList --webUrl https://contoso.sharepoint.com/sites/project-x
```

Extract a site script from an existing SharePoint list with specified id located in the specified site.

```sh
m365 spo list sitescript get --listId cc27a922-8224-4296-90a5-ebbc54da2e85 --webUrl https://contoso.sharepoint.com/sites/project-x
```

Extract a site script from an existing SharePoint list with specified server relative url located in the specified site.

```sh
m365 spo list sitescript get --listUrl 'sites/project-x/Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

Extract a site script from an existing SharePoint list with specified site-relative URL located in the specified site.

```sh
m365 spo list sitescript get --listUrl 'Shared Documents' --webUrl https://contoso.sharepoint.com/sites/project-x
```

