# spo list sitescript get

Extracts a site script from a SharePoint list

## Usage

```sh
m365 spo list sitescript get [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site where the list to extract the site script from is located

`-l, --listId [listId]`
: ID of the list to extract the site script from. Specify either `listId`, `title` or `listTitle` but not multiple.

`-t, --listTitle [listTitle]`
: Title of the list to extract the site script from. Specify either `listId`, `title` or `listTitle` but not multiple.

`--listUrl [listUrl]`
: Server- or site-relative URL of the list. Specify either `listId`, `title` or `listTitle` but not multiple.

--8<-- "docs/cmd/_global.md"

## Examples

Extract a site script from an existing SharePoint list with title _ContosoList_ located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle ContosoList
```

Extract a site script from an existing SharePoint list with id _cc27a922-8224-4296-90a5-ebbc54da2e85_
located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listId cc27a922-8224-4296-90a5-ebbc54da2e85
```

Extract a site script from an existing SharePoint list with server relative url _sites/project-x/Documents_
located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'sites/project-x/Documents'
```

Extract a site script from an existing SharePoint list with site-relative URL _Shared Documents_
located in site _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo list sitescript get --webUrl https://contoso.sharepoint.com/sites/project-x --listUrl 'Shared Documents'
```

