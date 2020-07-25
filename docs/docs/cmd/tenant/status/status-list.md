# tenant status list

Gets health status of the different services in Microsoft 365

## Usage

```sh
m365 tenant status list [options]
```

## Options

`-h, --help`
: output usage information

`-w, --workload [workload]`
: Retrieve service status for the specified service. If not provided, will list the current service status of all services

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json|text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

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
