# spo tenant applicationcustomizer remove

Remove an application customizer that is installed tenant wide.

## Usage

```sh
spo tenant applicationcustomizer remove [options]
```

## Options

`-t, --title [title]`
: The title of the Application Customizer. Specify either `title`, `id`, or `clientSideComponentId`.

`-i, --id [id]`
: The id of the Application Customizer. Specify either `title`, `id`, or `clientSideComponentId`.

`-c, --clientSideComponentId  [clientSideComponentId]`
: The Client Side Component Id (GUID) of the application customizer. Specify either `title`, `id`, or `clientSideComponentId`.

`--confirm`
: Don't prompt for confirmation.

--8<-- "docs/cmd/_global.md"

## Examples

Removes an application customizer by title.

```sh
m365 spo tenant applicationcustomizer remove --title "Some customizer"
```

Removes an application customizer by id.

```sh
m365 spo tenant applicationcustomizer remove --id 3
```

Removes an application customizer by clientSideComponentId.

```sh
m365 spo tenant applicationcustomizer remove --clientSideComponentId "7096cded-b83d-4eab-96f0-df477ed7c0bc"
```

## Response

The command won't return a response on success.
