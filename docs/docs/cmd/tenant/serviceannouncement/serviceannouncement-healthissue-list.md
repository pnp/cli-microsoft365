# tenant serviceannouncement healthissue list

Gets all service health issues for the tenant.

## Usage

```sh
m365 tenant serviceannouncement healthissue list [options]
```

## Options

`-s, --service [service]`
: Retrieve service health issues for the particular service. If not provided, retrieves health issues for all services

--8<-- "docs/cmd/\_global.md"

## Examples

Get service health issues of all services in Microsoft 365

```sh
m365 tenant serviceannouncement healthissue list
```

Get service health issues for Microsoft Forms

```sh
m365 tenant serviceannouncement healthissue list --service "Microsoft Forms"
```

## More information

- List serviceAnnouncement issues: [https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-issues](https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-issues)
