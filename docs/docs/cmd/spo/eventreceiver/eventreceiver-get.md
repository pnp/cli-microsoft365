# spo eventreceiver get

Retrieves specific event receiver for the specified web, site or list by event receiver name or id.

## Usage

```sh
m365 spo eventreceiver get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the web for which to retrieve the event receivers.

`--listTitle [listTitle]`
: The title of the list for which to retrieve the event receivers, _if the event receivers should be retrieved from a list_.
Specify either `listTitle`, `listId` or `listUrl`.

`--listId [listId]`
: The id of the list for which to retrieve the event receivers, _if the event receivers should be retrieved from a list_.
Specify either `listTitle`, `listId` or `listUrl`.

`--listUrl [listUrl]`
: The url of the list for which to retrieve the event receivers, _if the event receivers should be retrieved from a list_.
Specify either `listTitle`, `listId` or `listUrl`.

`-n, --name [name]`
The name of the event receiver to retrieve. Specify either `name` or `id` but not both.

`-i, --id [id]`
The id of the event receiver to retrieve. Specify either `name` or `id` but not both.

`-s, --scope [scope]`
: The scope of which to retrieve the Event Receivers.
Can be either "site" or "web". Defaults to "web". Only applicable when not specifying any of the list properties.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve event receivers in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --name 'PnP Test Receiver'
```

Retrieve event receivers in site _<https://contoso.sharepoint.com/sites/contoso-sales>_ with id _c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --scope site --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Retrieve event receivers for list with title _Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events --name 'PnP Test Receiver'
```

Retrieve event receivers for list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with id _c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '202b8199-b9de-43fd-9737-7f213f51c991' --id c5a6444a-9c7f-4a0d-9e29-fc6fe30e34ec
```

Retrieve event receivers for list with url _/sites/contoso-sales/lists/Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_ with name _PnP Test Receiver_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/sites/contoso-sales/lists/Events' --name 'PnP Test Receiver'
```
