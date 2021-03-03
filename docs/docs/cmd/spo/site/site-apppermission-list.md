# spo site apppermission list

Lists application permissions for a site

## Usage

```sh
m365 spo site apppermission list [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site collection to retrieve information for

`-i, --appId [appId]`
: Id of the application to filter by

`-n, --appDisplayName [appDisplayName]`
: Display name of the application to filter by

--8<-- "docs/cmd/_global.md"

## Remarks

To filter by an app, pass in either `appId` or `appDisplayName` not both

## Examples

Return list of application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection.

```sh
m365 spo site apppermission list --siteUrl https://contoso.sharepoint.com/sites/project-x
```

Return list of application permissions for the _https://contoso.sharepoint.com/sites/project-x_ site collection and filter by an application called Foo

```sh
m365 spo site apppermission list --siteUrl https://contoso.sharepoint.com/sites/project-x --appDisplayName Foo
```
