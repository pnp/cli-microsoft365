# tenant status list

Gets health status of the different services in Microsoft 365

## Usage

```sh
m365 tenant status list [options]
```

## Options

`-w, --workload [workload]`
: Retrieve service status for the specified service. If not provided, will list the current service status of all services

--8<-- "docs/cmd/_global.md"

## Examples

Gets health status of the different services in Microsoft 365

```sh
m365 tenant status list
```

Gets health status for SharePoint Online

```sh
m365 tenant status list --workload "SharePoint"
```

## More information

- Microsoft 365 Service Communications API reference: [https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-current-status](https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-current-status)
