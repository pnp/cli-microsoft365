# tenant service announcement health list

Gets the health report of all subscribed services for a tenant

## Usage

```sh
m365 tenant serviceannouncement health list [options]
```

## Options

`-i, --issues`
: Return the collection of issues that happened on the service, with detailed information for each issue. Is only returned in JSON output mode.

--8<-- "docs/cmd/_global.md"

## Examples

Get the health report of all subscribed services for a tenant

```sh
m365 tenant serviceannouncement health list
```

Get the health report of all subscribed services for a tenant including the issues that happend on each service

```sh
m365 tenant serviceannouncement health list --issues
```
