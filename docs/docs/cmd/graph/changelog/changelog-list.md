# graph changelog list

Gets an overview of specific API-level changes in Microsoft Graph v1.0 and beta

## Usage

```sh
m365 graph changelog list [options]
```

## Options

`-v, --versions [versions]`
: Comma-separated list of versions to show changes for. `Beta, v1.0`. When no version is selected all versions are returned.

`-c, --changeType [changeType]`
: Change type to show changes for. `Addition, Change, Deletion`. When no changeType is selected all change types are returned.

`-s, --services [services]`
: Comma-separated list of services to show changes for. `Applications, Calendar, Change notifications, Cloud communications, Compliance, Cross-device experiences, Customer booking, Device and app management, Education, Files, Financials, Groups, Identity and access, Mail, Notes, Notifications, People and workplace intelligence, Personal contacts, Reports, Search, Security, Sites and lists, Tasks and plans, Teamwork, To-do tasks, Users, Workbooks and charts`. When no service is selected all services are returned.

`--startDate [startDate]`
: The startdate used to query for changes. Supported date format is `YYYY-MM-DD`. When no date is specified all changes are returned.

`--endDate [endDate]`
: The enddate used to query for changes. Supported date format is `YYYY-MM-DD`. When no date is specified all changes are returned.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Get all changes within Microsoft Graph.

```sh
m365 graph changelog list
```

Get all changes within Microsoft Graph for the services _Groups_ and _Users_.

```sh
m365 graph changelog list --services 'Groups,Users'
```

Get all changes within Microsoft Graph that happend between _2021-01-01_ and _2021-05-01_.

```sh
m365 graph changelog list --startDate '2021-01-01' --endDate '2021-05-01'
```
