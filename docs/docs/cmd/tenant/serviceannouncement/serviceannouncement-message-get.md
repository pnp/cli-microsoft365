# tenant serviceannouncement message get

Retrieves a specified service update message for the tenant

## Usage

```sh
m365 tenant serviceannouncement message get [options]
```

## Options

`-i, --id <id>`
: The ID of the service update message.

--8<-- "docs/cmd/_global.md"

## Examples

Get service update message with ID MC001337

```sh
m365 tenant serviceannouncement message get --id MC001337
```

## More information

- Microsoft Graph REST API reference: [https://docs.microsoft.com/en-us/graph/api/serviceupdatemessage-get](https://docs.microsoft.com/en-us/graph/api/serviceupdatemessage-get)
