# spo site list

Lists classic sites of the given type

## Usage

```sh
spo site list classic [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--type [type]`|type of classic sites to list. Allowed values `STS#0|BLOG#0|BDR#0|DEV#0|OFFILE#1|EHS#1|BICenterSite#0|SRCHCEN#0|BLANKINTERNET#0|BLANKINTERNETCONTAINER#0|ENTERWIKI#0|PROJECTSITE#0|PRODUCTCATALOG#0|COMMUNITY#0|COMMUNITYPORTAL#0|SRCHCENTERLITE#0|visprus#0|GROUP#0|SITEPAGEPUBLISHING#0`
`-f, --filter [filter]`|filter to apply when retrieving sites
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To list classic sites, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to list the available sites, you will get an error.

Using the `-f, --filter` option you can specify which sites you want to retrieve. For example, to get sites with _project_ in their URL, use `Url -like 'project'` as the filter.

Using the `--includeOneDriveSites` option, you can include OneDrive sites in your request. If you don't use the option, it will leave the OneDrive sites out of the request.

## Examples

List all classic sites in the currently connected tenant

```sh
spo site list classic
```

List all classic team sites in the currently connected tenant including OneDrive sites

```sh
spo site list classic --includeOneDriveSites
```

List all classic team sites in the currently connected tenant

```sh
spo site list classic --type STS#0
```

List all classic project sites that contain _project_ in the URL

```sh
spo site list classic --type PROJECTSITE#0 --filter "Url -like 'project'"
```