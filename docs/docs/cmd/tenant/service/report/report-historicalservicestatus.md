# tenant service report historicalservicestatus

Gets the historical service status of the Office 365 Services of the last 7 days from the Office 365 Management API

## Usage

```sh
m365 tenant service report historicalservicestatus [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-w, --workload [workload]`|Allows retrieval of the historical service status of only one particular service. If not provided, the historical service status of all services will be returned.
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the historical service status of the Office 365 Services of the last 7 days

```sh
m365 tenant service report historicalservicestatus
```

Gets the historical service status of the Office 365 Services of the last 7 days for SharePoint Online

```sh
m365 tenant service report historicalservicestatus --workload "SharePoint"
```

## More information

- Microsoft 365 Service Communications API reference: [https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-historical-status](https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-historical-status)
