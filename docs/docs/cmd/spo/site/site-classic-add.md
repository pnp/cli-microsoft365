# spo site classic add

Creates new classic site.

## Usage

```sh
m365 spo site classic add [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The absolute site url

`-t, --title <title>`
: The site title

`--owner <owner>`
: The account name of the site owner

`-z, --timeZone <timeZone>`
: Integer representing time zone to use for the site

`-l, --lcid [lcid]`
: Integer representing time zone to use for the site

`-w, --webTemplate [webTemplate]`
: Template to use for creating the site. Default `STS#0`

`--resourceQuota [resourceQuota]`
: The quota for this site collection in Sandboxed Solutions units. Default `0`

`--resourceQuotaWarningLevel [resourceQuotaWarningLevel]`
: The warning level for the resource quota. Default `0`

`--storageQuota [storageQuota]`
: The storage quota for this site collection in megabytes. Default `100`

`--storageQuotaWarningLevel [storageQuotaWarningLevel]`
: The warning level for the storage quota in megabytes. Default `100`

`--removeDeletedSite`
: Set, to remove existing deleted site with the same URL from the Recycle Bin

`--wait`
: Wait for the site to be provisioned before completing the command

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

!!! important
    This command is deprecated. Please use [spo site add](./site-add.md) instead.

## Remarks

Using the `-z, --timeZone` option you have to specify the time zone of the site. For more information about the valid values see [https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx](https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx).

The `-l, --lcid` option denotes the language of the site. For more information see Locale IDs Assigned by Microsoft: [https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx](https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx).

The value of the `--resourceQuota` option must not exceed the company's aggregate available Sandboxed Solutions quota. For more information, see Resource Usage Limits on Sandboxed Solutions in SharePoint 2010: [http://msdn.microsoft.com/en-us/library/gg615462.aspx](http://msdn.microsoft.com/en-us/library/gg615462.aspx).

The value of the `--resourceQuotaWarningLevel` option must not exceed the value of the `--resourceQuota` option.

The value of the `--storageQuota` option must not exceed the company's available quota.

The value of the `--storageQuotaWarningLevel` option must not exceed the the value of the `--storageQuota` option.

If you try to create a site with the same URL as a site that has been previously moved to the recycle bin, you will get an error. To avoid this error, you can use the `--removeDeletedSite` option. Prior to creating the site, the spo site classic add command will check if the site with the specified URL has been previously moved to the recycle bin and if so, will remove it. Because removing sites from the recycle bin might take a moment, it should be used in conjunction with the `--wait` option so that the new site is not created before the old site is fully removed.

Deleting and creating classic site collections is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully provisioned, you should use the `--wait` flag. When using this flag, the spo site classic add command will keep running until it received confirmation from Microsoft 365 that the site has been fully provisioned.

## Examples

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. Don't wait for the site provisioning to complete

```sh
m365 spo site classic add --url https://contoso.sharepoint.com/sites/team --title Team --owner admin@contoso.onmicrosoft.com --timeZone 4
```

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. Wait for the site provisioning to complete

```sh
m365 spo site classic add --url https://contoso.sharepoint.com/sites/team --title Team --owner admin@contoso.onmicrosoft.com --timeZone 4 --webTemplate STS#0 --wait
```

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. If a site with the same URL is in the recycle bin, delete it. Wait for the site provisioning to complete

```sh
m365 spo site classic add --url https://contoso.sharepoint.com/sites/team --title Team --owner admin@contoso.onmicrosoft.com --timeZone 4 --webTemplate STS#0 --removeDeletedSite --wait
```