# tenant service message list

Gets the service messages regarding services in Office 365 from the Office 365 Management API

## Usage

```sh
tenant service message list [options]
```

## Options

Option|Description
------|-----------
`-w, --workload [workload]`|Allows retrieval of the service messages for only one particular service. If not provided, the service messages of all services will be returned.
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get service messages of all services in Microsoft 365

```sh
tenant service message list
```

Get service messages of only one particular service in Microsoft 365

```sh
tenant service message list -w "Exchange Online"
```

## More information

- Microsoft 365 Service Communications API reference: [https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages](https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages)