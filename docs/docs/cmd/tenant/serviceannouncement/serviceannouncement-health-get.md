# tenant service announcement health get

Get the health report of a specified service for a tenant

## Usage

```sh
m365 tenant serviceannouncement health get [options]
```

## Options

`-s, --serviceName <serviceName>`
: The service name to retrieve the health report for.

`-i, --issues`
: Return the collection of issues that happened on the service, with detailed information for each issue. Is only returned in JSON output mode.

--8<-- "docs/cmd/\_global.md"

## Examples

Get the health report for the service _Exchange Online_

```sh
m365 tenant serviceannouncement health get --serviceName "Exchange Online"
```

Get the health report for the service _Exchange Online_ including the issues of the service

```sh
m365 tenant serviceannouncement health get --serviceName "Exchange Online" --issues
```
