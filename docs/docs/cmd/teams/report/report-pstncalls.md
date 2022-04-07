# teams report pstncalls

Get details about PSTN calls made within a given time period

## Usage

```sh
m365 teams report pstncalls [options]
```

## Options

`--fromDateTime <fromDateTime>`
: The start of time range to query. UTC, inclusive

`--toDateTime [toDateTime]`
: The end time range to query. UTC, inclusive. Defaults to today if omitted

--8<-- "docs/cmd/_global.md"

## Remarks

This command only works with app-only permissions. You will need to create your own Azure AD app with `CallRecords.Read.All` permission assigned. Instructions on how to create your own Azure AD app can be found at [Using your own Azure AD identity](../../../user-guide/using-own-identity.md)

The difference between `fromDateTime` and `toDateTime` cannot exceed a period of 90 days

## Examples

Get details about PSTN calls made between 2020-10-31 and today

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in text format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output text > "pstncalls.txt"
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in json format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output json > "pstncalls.json"
```

## More information

- List PSTN calls: [https://docs.microsoft.com/en-us/graph/api/callrecords-callrecord-getpstncalls?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/callrecords-callrecord-getpstncalls?view=graph-rest-1.0)
