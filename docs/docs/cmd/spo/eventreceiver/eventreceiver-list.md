# spo eventreceiver list

Retrieves event receivers for the specified web, site or list.

## Usage

```sh
m365 spo eventreceiver list [options]
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

`-s, --scope [scope]`
: The scope of which to retrieve the Event Receivers.
Can be either "site" or "web". Defaults to "web". Only applicable when not specifying any of the list properties.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieves event receivers in web _<https://contoso.sharepoint.com/sites/contoso-sales>_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales
```

Retrieves event receivers in site _<https://contoso.sharepoint.com/sites/contoso-sales>_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --scope site
```

Retrieves event receivers for list with title _Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listTitle Events
```

Retrieves event receivers for list with ID _202b8199-b9de-43fd-9737-7f213f51c991_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listId '202b8199-b9de-43fd-9737-7f213f51c991'
```

Retrieves event receivers for list with url _/sites/contoso-sales/lists/Events_ in web _<https://contoso.sharepoint.com/sites/contoso-sales>_.

```sh
m365 spo eventreceiver list --webUrl https://contoso.sharepoint.com/sites/contoso-sales --listUrl '/sites/contoso-sales/lists/Events'
```
