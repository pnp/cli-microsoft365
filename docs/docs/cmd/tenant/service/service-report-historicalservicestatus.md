# tenant service report historicalservicestatus

Gets the historical service status of Microsoft 365 Services of the last 7 days

## Usage

```sh
m365 tenant service report historicalservicestatus [options]
```

## Options

Option|Description
------|-----------
`-h, --help`|output usage information
`-w, --workload [workload]`|Retrieve the historical service status for the particular service. If not provided, the historical service status of all services will be returned.
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

To get the name of the particular workload for use with the workload option, execute `m365 tenant service report historicalservicestatus --output json` and get the value of the `Workload` property for the particular service.

## Examples

Gets the historical service status of Microsoft 365 Services of the last 7 days

```sh
m365 tenant service report historicalservicestatus
```

Gets the historical service status of Microsoft Teams for the last 7 days

```sh
m365 tenant service report historicalservicestatus --workload "microsoftteams"
```

## More information

- Microsoft 365 Service Communications API reference: [https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-historical-status](https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference#get-historical-status)
