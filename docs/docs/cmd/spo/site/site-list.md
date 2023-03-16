# spo site list

Lists modern sites of the given type

## Usage

```sh
m365 spo site list [options]
```

## Options

`-t, --type [type]`
: convenience option for type of sites to list. Allowed values are `TeamSite,CommunicationSite`.

`--webTemplate [webTemplate]`
: type of sites to list. To be used with values like `GROUP#0` and `SITEPAGEPUBLISHING#0`. Specify either `type` or `webTemplate`, but not both.  

`-f, --filter [filter]`
: filter to apply when retrieving sites

`--includeOneDriveSites`
: use this switch to include OneDrive sites in the result when retrieving sites. Do not specify the `type` or `webTemplate` options when using this.

`--deleted`
: use this switch to only return deleted sites

--8<-- "docs/cmd/_global.md"

## Remarks

Using the `-f, --filter` option you can specify which sites you want to retrieve. For example, to get sites with _project_ in their URL, use `Url -like 'project'` as the filter.

When using the text output type, the command lists only the values of the `Title`, and `Url` properties of the site. When setting the output type to JSON, all available properties are included in the command output.

!!! important
    To use this command you have to have permissions to access the tenant admin site.
    
## Examples

List all sites in the currently connected tenant

```sh
m365 spo site list
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
m365 spo site list --includeOneDriveSites
```

List all deleted sites in the tenant you're logged in to

```sh
m365 spo site list --deleted
```
