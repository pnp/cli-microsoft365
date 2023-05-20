# purview retentionlabel remove

Delete a retention label

## Usage

```sh
m365 purview retentionlabel remove [options]
```

## Options

`-i, --id <id>`
: The Id of the retention label.

`--confirm`
: Don't prompt for confirming to remove the label.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

!!! attention
    This command currently only supports delegated permissions.

## Examples

Delete a retention label

```sh
m365 purview retentionlabel remove --id 'e554d69c-0992-4f9b-8a66-fca3c4d9c531'
```

## Response

The command won't return a response on success.
