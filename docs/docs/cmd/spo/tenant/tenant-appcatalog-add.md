# spo tenant appcatalog add

Creates new tenant app catalog site

## Usage

```sh
m365 spo tenant appcatalog add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The absolute site url

`--owner <owner>`
: The account name of the site owner

`-z, --timeZone <timeZone>`
: Integer representing time zone to use for the site

`--wait`
: Wait for the site to be provisioned before completing the command

`--force`
: Force creating a new app catalog site if one already exists

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

If there is an app catalog registered in your tenant, creating a new app catalog using this command will fail, unless you use the `force` option.

If you use the `force` option, and either the app catalog or the site at the specified URL exists (or both), this command will delete both sites bypassing the recycle bin.

Creating an app catalog site might take a while. If you need to wait for the site to be created before continuing, use the `wait` option.

## Examples

Creates new app catalog. Will fail if another app catalog or site at the specified URL exists

```sh
m365 spo tenant appcatalog add --url https://contoso.sharepoint.com/sites/apps --owner admin@contoso.com --timeZone 4
```

Creates new app catalog and waits for completion. Will fail if another app catalog or site at the specified URL exists

```sh
m365 spo tenant appcatalog add --url https://contoso.sharepoint.com/sites/apps --owner admin@contoso.com --timeZone 4 --wait
```

Creates new app catalog and deletes existing app catalog and existing site at the specified URL

```sh
m365 spo tenant appcatalog add --url https://contoso.sharepoint.com/sites/apps --owner admin@contoso.com --timeZone 4 --force
```
