# spo site list

Lists modern sites of the given type

## Usage

```sh
m365 spo site list [options]
```

## Options

`-t, --type [type]`
: type of sites to list. Allowed values are `TeamSite,CommunicationSite,All`. The default value is `TeamSite`. 

`--webTemplate [webTemplate]`
: types of sites to list. To be used with values like `GROUP#0` and `SITEPAGEPUBLISHING#0`. Specify either `type` or `webTemplate`, but not both.  

`-f, --filter [filter]`
: filter to apply when retrieving sites

`--includeOneDriveSites`
: use this switch to include OneDrive sites in the result when retrieving sites. Can only be used in combination with `type` All.

`--deleted`
: use this switch to only return deleted sites

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-f, --filter` option you can specify which sites you want to retrieve. For example, to get sites with _project_ in their URL, use `Url -like 'project'` as the filter.

When using the text output type (default), the command lists only the values of the `Title`, and `Url` properties of the site. When setting the output type to JSON, all available properties are included in the command output.

## Examples

List all sites in the currently connected tenant

```sh
m365 spo site list --type All
```

List all group connected team sites in the currently connected tenant

```sh
m365 spo site list --type TeamSite
```

List all communication sites in the currently connected tenant

```sh
m365 spo site list --type CommunicationSite
```

List all group connected team sites that contain _project_ in the URL

```sh
m365 spo site list --type TeamSite --filter "Url -like 'project'"
```

List all sites in the currently connected tenant including OneDrive sites

```sh
m365 spo site list --type All --includeOneDriveSites
```

List all deleted sites in the tenant you're logged in to

```sh
m365 spo site list --deleted
```
