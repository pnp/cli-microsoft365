# tenant serviceannouncement message list

Gets all service update messages that exist for the tenant.

## Usage

```sh
m365 tenant serviceannouncement message list [options]
```

## Options

`-s, --service [service]`
: Retrieve service update messages for the particular service. If not provided, retrieves messages for all services

--8<-- "docs/cmd/_global.md"

## Examples

Get service update messages of all services in Microsoft 365

```sh
m365 tenant serviceannouncement message list
```

Get service update messages for Microsoft Teams

```sh
m365 tenant serviceannouncement message list --service "Microsoft Teams"
```

## More information

- List serviceAnnouncement messages: [https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-messages](https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-messages)