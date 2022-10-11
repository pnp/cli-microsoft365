# spo eventreceiver remove

Removes event receivers for the specified web, site, or list.

## Usage

```sh
m365 spo eventreceiver remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the web from which to remove the event receiver.

`--listTitle [listTitle]`
: The title of the list from which to remove the event receiver, _if the event receiver should be removed from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: The id of the list from which to remove the event receiver, _if the event receiver should be retrieved from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: The url of the list from which to remove the event receiver, _if the event receiver should be retrieved from a list_. Specify either `listTitle`, `listId` or `listUrl`.

`-n, --name [name]`
The name of the event receiver to remove. Specify either `name` or `id` but not both.

`-i, --id [id]`
The id of the event receiver to remove. Specify either `name` or `id` but not both.

`-s, --scope [scope]`
: The scope of which to remove the Event Receiver.
Can be either "site" or "web". Defaults to "web". Only applicable when not specifying any of the list properties.

`--confirm`
: Don't prompt for confirming removing the field

--8<-- "docs/cmd/_global.md"

## Examples

Remove event receiver in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'PnP Test Receiver'
```

Remove event receiver in site _<https://contoso.sharepoint.com/sites/contoso-sales>_ with id _c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec_.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --scope site --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Remove event receiver for list with title _Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --name 'PnP Test Receiver'
```

Remove event receiver for list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with id _c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec_.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '202b8199-b9de-43fd-9737-7f213f51c991' --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Remove event receiver for list with url _/sites/contoso-sales/lists/Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver remove --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/sites/contoso-sales/lists/Events' --name 'PnP Test Receiver'
```
