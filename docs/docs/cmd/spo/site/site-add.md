# spo site add

Creates new SharePoint Online site

## Usage

```sh
m365 spo site add [options]
```

## Options

`-h, --help`
: output usage information

`--type [type]`
: Type of sites to add. Allowed values `TeamSite,CommunicationSite,ClassicSite`, default `TeamSite`

`-t, --title <title>`
: Site title

`-a, --alias [alias]`
: Site alias, used in the URL and in the team site group e-mail (applies to type TeamSite)

`-u, --url [url]`
: Site URL  (applies to type CommunicationSite, ClassicSite)

`-z, --timeZone [timeZone]`
: Integer representing time zone to use for the site (applies to type ClassicSite)

`-d, --description [description]`
: Site description

`-l, --lcid [lcid]`
: Site language in the LCID format, eg. _1033_ for _en-US_. See [SharePoint documentation](https://support.microsoft.com/en-us/office/languages-supported-by-sharepoint-dfbf3652-2902-4809-be21-9080b6512fff) for the list of supported languages

`--owners [owners]`
: Comma-separated list of users to set as site owners (applies to type TeamSite, ClassicSite)

`--isPublic`
: Determines if the associated group is public or not (applies to type TeamSite)

`-c, --classification [classification]`
: Site classification (applies to type TeamSite, CommunicationSite)

`--siteDesign [siteDesign]`
: Type of communication site to create. Allowed values `Topic,Showcase,Blank`, default `Topic`. When creating a communication site, specify either `siteDesign` or `siteDesignId` (applies to type CommunicationSite)

`--siteDesignId [siteDesignId]`
: Id of the custom site design to use to create the site. When creating a communication site, specify either `siteDesign` or `siteDesignId` (applies to type CommunicationSite)

`--allowFileSharingForGuestUsers`
: (deprecated. Use `shareByEmailEnabled` instead) Determines whether it's allowed to share file with guests (applies to type CommunicationSite)

`--shareByEmailEnabled`
: Determines whether it's allowed to share file with guests (applies to type CommunicationSite)

`-w, --webTemplate [webTemplate]`
: Template to use for creating the site. Default `STS#0`  (applies to type ClassicSite)

`--resourceQuota [resourceQuota]`
: The quota for this site collection in Sandboxed Solutions units. Default `0`  (applies to type ClassicSite)

`--resourceQuotaWarningLevel [resourceQuotaWarningLevel]`
: The warning level for the resource quota. Default `0`  (applies to type ClassicSite)

`--storageQuota [storageQuota]`
: The storage quota for this site collection in megabytes. Default `100`  (applies to type ClassicSite)

`--storageQuotaWarningLevel [storageQuotaWarningLevel]`
: The warning level for the storage quota in megabytes. Default `100`  (applies to type ClassicSite)

`--removeDeletedSite`
: Set, to remove existing deleted site with the same URL from the Recycle Bin  (applies to type ClassicSite)

`--wait`
: Wait for the site to be provisioned before completing the command  (applies to type ClassicSite)

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks for classic sites

Using the `-z, --timeZone` option you have to specify the time zone of the site. For more information about the valid values see [https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx](https://msdn.microsoft.com/library/microsoft.sharepoint.spregionalsettings.timezones.aspx).

The value of the `--resourceQuota` option must not exceed the company's aggregate available Sandboxed Solutions quota. For more information, see Resource Usage Limits on Sandboxed Solutions in SharePoint 2010: [http://msdn.microsoft.com/en-us/library/gg615462.aspx](http://msdn.microsoft.com/en-us/library/gg615462.aspx).

The value of the `--resourceQuotaWarningLevel` option must not exceed the value of the `--resourceQuota` option.

The value of the `--storageQuota` option must not exceed the company's available quota.

The value of the `--storageQuotaWarningLevel` option must not exceed the the value of the `--storageQuota` option.

If you try to create a site with the same URL as a site that has been previously moved to the recycle bin, you will get an error. To avoid this error, you can use the `--removeDeletedSite` option. Prior to creating the site, the spo site classic add command will check if the site with the specified URL has been previously moved to the recycle bin and if so, will remove it. Because removing sites from the recycle bin might take a moment, it should be used in conjunction with the `--wait` option so that the new site is not created before the old site is fully removed.

Deleting and creating classic site collections is by default asynchronous and depending on the current state of Office 365, might take up to few minutes. If you're building a script with steps that require the site to be fully provisioned, you should use the `--wait` flag. When using this flag, the spo site classic add command will keep running until it received confirmation from Office 365 that the site has been fully provisioned.

## Remarks for modern sites

The `--owners` option is mandatory for creating `CommunicationSite` sites with app-only permissions.

When trying to create a team site using app-only permissions, you will get an _Insufficient privileges to complete the operation._ error. As a workaround, you can use the [`aad o365group add`](../../aad/o365group/o365group-add.md) command, followed by [`spo site set`](./site-set.md) to further configure the Team site.

## Examples

Create modern team site with private group

```sh
m365 spo site add --alias team1 --title "Team 1"
```

Create modern team site with description and classification

```sh
m365 spo site add --type TeamSite --alias team1 --title "Team 1" --description "Site of team 1" --classification LBI
```

Create modern team site with public group

```sh
m365 spo site add --type TeamSite --alias team1 --title "Team 1" --isPublic
```

Create modern team site using the Dutch language

```sh
m365 spo site add --alias team1 --title "Team 1" --lcid 1043
```

Create modern team site with the specified users as owners

```sh
m365 spo site add --alias team1 --title "Team 1" --owners "steve@contoso.com, bob@contoso.com"
```

Create communication site using the Topic design

```sh
m365 spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing
```

Create communication site using app-only permissions

```sh
m365 spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --owners "john.smith@contoso.com"
```

Create communication site using the Showcase design

```sh
m365 spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesign Showcase
```

Create communication site using a custom site design

```sh
m365 spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --siteDesignId 99f410fe-dd79-4b9d-8531-f2270c9c621c
```

Create communication site using the Blank design with description and classification

```sh
m365 spo site add --type CommunicationSite --url https://contoso.sharepoint.com/sites/marketing --title Marketing --description Site of the marketing department --classification MBI --siteDesign Blank
```

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. Don't wait for the site provisioning to complete

```sh
m365 spo site add --type ClassicSite --url https://contoso.sharepoint.com/sites/team --title Team --owners admin@contoso.onmicrosoft.com --timeZone 4
```

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. Wait for the site provisioning to complete

```sh
m365 spo site add --type ClassicSite --url https://contoso.sharepoint.com/sites/team --title Team --owners admin@contoso.onmicrosoft.com --timeZone 4 --webTemplate STS#0 --wait
```

Create new classic site collection using the Team site template. Set time zone to `UTC+01:00`. If a site with the same URL is in the recycle bin, delete it. Wait for the site provisioning to complete

```sh
m365 spo site add --type ClassicSite --url https://contoso.sharepoint.com/sites/team --title Team --owners admin@contoso.onmicrosoft.com --timeZone 4 --webTemplate STS#0 --removeDeletedSite --wait
```

## More information

- Creating SharePoint Communication Site using REST: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/communication-site-creation-rest](https://docs.microsoft.com/en-us/sharepoint/dev/apis/communication-site-creation-rest)