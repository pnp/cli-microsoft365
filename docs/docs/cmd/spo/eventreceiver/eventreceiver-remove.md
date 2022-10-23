# spo eventreceiver remove

Removes event receivers for the specified web, site, or list.

## Usage

```sh
m365 spo eventreceiver remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the web.

`--listTitle [listTitle]`
: The title of the list, _if the event receiver should be removed from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: The id of the list, _if the event receiver should be retrieved from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: The url of the list, _if the event receiver should be retrieved from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`-n, --name [name]`
: The name. Specify either `name` or `id` but not both.

`-i, --id [id]`
: The id. Specify either `name` or `id` but not both.

`-s, --scope [scope]`
: The scope. Can be either "site" or "web". Defaults to "web". Only applicable when not specifying any of the list properties.

`--confirm`
: Don't prompt for confirming removing the event receiver

--8<-- "docs/cmd/_global.md"

## Examples

Remove event receiver in a specific web by name.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'PnP Test Receiver'
```

Remove event receiver in a specific site by id.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --scope site --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Remove event receiver in a specific list retrieved by title by name.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --name 'PnP Test Receiver'
```

Remove event receiver in a specific list retrieved by list id by id.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '202b8199-b9de-43fd-9737-7f213f51c991' --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Remove event receiver in a specific list retrieved by list url by name.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/sites/contoso-sales/lists/Events' --name 'PnP Test Receiver'
```
