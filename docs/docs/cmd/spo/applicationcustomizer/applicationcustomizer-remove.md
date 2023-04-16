# spo applicationcustomizer remove

Remove an application customizer from a site

## Usage

```sh
m365 spo applicationcustomizer remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site

`-t, --title <title>`
: The title of the application customizer

`-i, --id <id>`
: ID of the application customizer (GUID)

`-c, --clientSideComponentId <clientSideComponentId>`
: Client-side component ID of the application customizer (GUID)

`-s, --scope [scope]`
: Scope of the application customizer. Allowed values: `Site`, `Web`, and `All`. Defaults to `All`

--8<-- "docs/cmd/_global.md"

## Remarks

If the command finds multiple application customizers with the specified title or clientSideComponentId, it will prompt you to disambiguate which customizer it should remove, listing the discovered IDs.

## Examples

Remove an application customizer by id

```sh
m365 spo applicationcustomizer remove --id 14125658-a9bc-4ddf-9c75-1b5767c9a337 --webUrl https://contoso.sharepoint.com/sites/sales
```

Remove an application customizer by title

```sh
m365 spo applicationcustomizer remove --title "Some customizer" --webUrl https://contoso.sharepoint.com/sites/sales
```

Remove an application customizer by clientSideComponentId

```sh
m365 spo applicationcustomizer remove --clientSideComponentId 7096cded-b83d-4eab-96f0-df477ed7c0bc --webUrl https://contoso.sharepoint.com/sites/sales
```

Remove an application customizer by its id without prompting for confirmation

```sh
m365 spo applicationcustomizer remove --id 14125658-a9bc-4ddf-9c75-1b5767c9a337 --webUrl https://contoso.sharepoint.com/sites/sales --confirm
```

Remove an application customizer from a site collection by its id without prompting for confirmation

```sh
m365 spo applicationcustomizer remove --id 14125658-a9bc-4ddf-9c75-1b5767c9a337 --webUrl https://contoso.sharepoint.com/sites/sales --confirm --scope Site
```

Remove an application customizer from a site by its id without prompting for confirmation

```sh
m365 spo applicationcustomizer remove --id 14125658-a9bc-4ddf-9c75-1b5767c9a337 --webUrl https://contoso.sharepoint.com/sites/sales --confirm --scope Web
```

## Response

The command won't return a response on success.
