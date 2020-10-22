# tenant report servicesusercounts

Gets the count of users by activity type and service.

## Usage

```sh
m365 tenant report servicesusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of users by activity type and service for the last week

```sh
m365 tenant report servicesusercounts --period D7
```

Gets the count of users by activity type and service for the last week and exports the report data in the specified path in text format

```sh
m365 tenant report servicesusercounts --period D7 --output text > "servicesusercounts.txt"
```

Gets the count of users by activity type and service for the last week and exports the report data in the specified path in json format

```sh
m365 tenant report servicesusercounts --period D7 --output json > "servicesusercounts.json"
```
