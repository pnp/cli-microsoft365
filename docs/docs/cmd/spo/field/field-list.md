# spo field list

Retrieves fields for a given list or site

## Usage

```sh
m365 spo field list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the fields are located

`-l, --listTitle [listTitle]`
: Title of the list where the fields are located. Specify only one of listTitle, listId or listUrl

`--listId [listId]`
: ID of the list where the fields are located. Specify only one of listTitle, listId or listUrl

`--listUrl [listUrl]`
: Server- or web-relative URL of the list where the fields are located. Specify only one of listTitle, listId or listUrl


--8<-- "docs/cmd/_global.md"

## Examples

Retrieves site columns located in site _https://contoso.sharepoint.com/sites/contoso-sales_

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales
```

Retrieves list columns located in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its title

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events
```

Retrieves list columns located in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its url

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl 'Lists/Events'
```