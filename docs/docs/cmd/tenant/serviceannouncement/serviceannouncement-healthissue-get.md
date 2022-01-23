# tenant serviceannouncement healthissue get

Gets a specified service health issue for tenant.

## Usage

```sh
m365 tenant serviceannouncement healthissue get [options]
```

## Options

`-i, --issueId <issueId>`
: The issue id to get details for

--8<-- "docs/cmd/_global.md"

## Examples

Get specified service health issue for tenant with issueId _MO226784_

```sh
m365 tenant serviceannouncement healthissue get --issueId MO226784
```

## More information

- Get serviceHealthIssue: [https://docs.microsoft.com/en-us/graph/api/servicehealthissue-get](https://docs.microsoft.com/en-us/graph/api/servicehealthissue-get)