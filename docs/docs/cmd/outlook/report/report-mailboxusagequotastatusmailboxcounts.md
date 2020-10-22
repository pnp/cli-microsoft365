# outlook report mailboxusagequotastatusmailboxcounts

Gets the count of user mailboxes in each quota category

## Usage

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of user mailboxes in each quota category for the last week

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7
```

Gets the count of user mailboxes in each quota category for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7 --output text > "mailboxusagequotastatusmailboxcounts.txt"
```

Gets the count of user mailboxes in each quota category for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7 --output json > "mailboxusagequotastatusmailboxcounts.json"
```
