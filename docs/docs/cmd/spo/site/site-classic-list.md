# spo site classic list

Lists classic sites of the given type

## Usage

```sh
m365 spo site classic list [options]
```

## Options

`-h, --help`
: output usage information

`-t, --webTemplate [type]`
: type of classic sites to list.

`-f, --filter [filter]`
: filter to apply when retrieving sites

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--includeOneDriveSites`
: option to include OneDrive sites or not.

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --webTemplate` option you can specify which sites you want to retrieve. For example, to get sites with the `STS#0` as their web template, use `--webTemplate STS#0` as the option.

Using the `-f, --filter` option you can specify which sites you want to retrieve. For example, to get sites with _project_ in their URL, use `Url -like 'project'` as the filter.

Using the `--includeOneDriveSites`option you can specify whether you want to retrieve OneDrive sites or not. For example, to retrieve OneDrive sites, use `--includeOneDriveSites` as the option.

## Examples

List all classic sites in the currently connected tenant

```sh
m365 spo site classic list
```

List all classic team sites in the currently connected tenant including OneDrive sites

```sh
m365 spo site classic list --includeOneDriveSites
```

List all classic team sites in the currently connected tenant

```sh
m365 spo site classic list --webTemplate STS#0
```

List all classic project sites that contain _project_ in the URL

```sh
m365 spo site classic list --webTemplate PROJECTSITE#0 --filter "Url -like 'project'"
```