# tenant auditlog report

Gets audit logs from the Office 365 Management API

## Usage

```sh
m365 tenant auditlog report [options]
```

## Options

`-c, --contentType <contentType>`
: Audit content type of logs to be retrieved, should be one of the following: `AzureActiveDirectory`, `Exchange`, `SharePoint`, `General`, `DLP`.

`-s, --startTime [startTime]`
: Start time of logs to be retrieved. Start time and end time must be less than or equal to 24 hours apart. Start time is mandatory if End time is specified.

`-e, --endTime [endTime]`
: End time of logs to be retrieved. Start time and end time must be less than or equal to 24 hours apart. If End time is not specified, command will assume the End time to be 24 hours from the specified Start time.

--8<-- "docs/cmd/_global.md"

## Remarks

By default, if `startTime` and `endTime` are not mentioned, then the content available in the last **24 hours** is returned. `startTime` and `endTime` must be less than or equal to **24 hours** apart, with the `startTime` prior to `endTime` and `startTime` no more than 7 days in the past.

If `endTime` is not specified, command will assume the `endTime` to be **24 hours** from the specified `startTime`.
`startTime` is mandatory if `endTime` is specified.

`DLP` audit log data is only available to users that have been granted “Read DLP sensitive data” permission. Otherwise you will get `Error: Request failed with status code 401`

## Examples

Gets audit logs from the Office 365 Management API for the `Exchange` content type.

```sh
m365 tenant auditlog report --contentType "Exchange"
```

Gets audit logs from the Office 365 Management API for the `Exchange` content type in the date range between `2020-12-13` and `2020-12-14`

```sh
m365 tenant auditlog report --contentType "Exchange" --startTime "2020-12-13" --endTime "2020-12-14"
```

Gets audit logs from the Office 365 Management API for the `Exchange` content type between `15:00` hours and `16:00` hours on `2020-12-13`

```sh
m365 tenant auditlog report --contentType "Exchange" --startTime "2020-12-13T15:00:00" --endTime "2020-12-13T16:00:00"
```

Gets audit logs from the Office 365 Management API for the `Exchange` content type between `23:00` hours on `2020-12-13` and `05:00` hours on `2020-12-14`

```sh
m365 tenant auditlog report --contentType "Exchange" --startTime "2020-12-13T23:00:00" --endTime "2020-12-14T05:00:00"
```

## More information

- Office 365 Management Activity API reference: [https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference](https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-management-activity-api-reference)
