# spo field list

Retrieves list of columns for specified list or site

## Usage

```sh
m365 spo field list [options]
```

## Options

`-u, --webUrl <webUrl>`
: Absolute URL of the site where the fields is located

`-t, --listTitle [listTitle]`
: Title of the list where the fields are located. Specify only one of listTitle, listId or listUrl

`-i --listId [listId]`
: ID of the list where the fields are located. Specify only one of listTitle, listId or listUrl

`--listUrl [listUrl]`
: Server- or web-relative URL of the list where the fields are located. Specify only one of listTitle, listId or listUrl

--8<-- "docs/cmd/_global.md"

## Examples

Retrieves site columns for _https://contoso.sharepoint.com/sites/contoso-sales_.

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales
```

Retrieves list columns for _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its title

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events
```

Retrieves list columns in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its id

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '202b8199-b9de-43fd-9737-7f213f51c991'
```

Retrieves list columns in site _https://contoso.sharepoint.com/sites/contoso-sales_. Retrieves the list by its url

```sh
m365 spo field list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/sites/contoso-sales/lists/Events'
```