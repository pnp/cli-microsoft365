# tenant service message list

Gets service messages Microsoft 365

## Usage

```sh
m365 tenant service message list [options]
```

## Options

`-h, --help`
: output usage information

`-w, --workload [workload]`
: Retrieve service messages for the particular workload. If not provided, retrieves messages for all workloads

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json|text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get service messages of all services in Microsoft 365

```sh
m365 tenant service message list
```

Get service messages for Microsoft Teams

```sh
m365 tenant service message list --workload microsoftteams
```

## More information

- Microsoft 365 Service Communications API reference: [https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages](https://docs.microsoft.com/office/office-365-management-api/office-365-service-communications-api-reference#get-messages)
