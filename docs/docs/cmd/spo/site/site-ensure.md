# spo site ensure

Ensures that the particular site collection exists and updates its properties if necessary.

The command is a combination of `spo web get`, `spo site add` and `spo site set`.

## Usage

```sh
m365 spo site ensure [options]
```

## Options

`-u, --url <url>`
: URL of the site to ensure

`--type [type]`
: Expected type of the site. Allowed values `TeamSite,CommunicationSite,ClassicSite`. If not specified, will not check the type of the site if one exists at the specified URL.

`-t, --title <title>`
: Site title

`-a, --alias [alias]`
: Site alias, used in the URL and in the team site group e-mail (applies to type `TeamSite`)

`-z, --timeZone [timeZone]`
: Integer representing time zone to use for the site (applies to type `ClassicSite`)

`-d, --description [description]`
: Site description

`-l, --lcid [lcid]`
: Site language in the LCID format, eg. _1033_ for _en-US_. See [SharePoint documentation](https://support.microsoft.com/en-us/office/languages-supported-by-sharepoint-dfbf3652-2902-4809-be21-9080b6512fff) for the list of supported languages

`--owners [owners]`
: Comma-separated list of users to set as site owners (applies to type TeamSite, ClassicSite)

`--isPublic`
: Determines if the associated group is public or not (applies to type `TeamSite`)

`-c, --classification [classification]`
: Site classification (applies to type `TeamSite`, `CommunicationSite`)

`--siteDesign [siteDesign]`
: Type of communication site to create. Allowed values `Topic,Showcase,Blank`, default `Topic`. Used only for creating site when one doesn't exist.

`--siteDesignId [siteDesignId]`
: Id of the custom site design to use to create the site. When creating a communication site, specify either `siteDesign` or `siteDesignId` (applies to type `CommunicationSite`)

`--shareByEmailEnabled`
: Determines whether it's allowed to share file with guests (applies to type `CommunicationSite`)

`-w, --webTemplate [webTemplate]`
: Template to use for creating the site. Default `STS#0`  (applies to type `ClassicSite`)

`--resourceQuota [resourceQuota]`
: The quota for this site collection in Sandboxed Solutions units. Default `0`  (applies to type `ClassicSite`)

`--resourceQuotaWarningLevel [resourceQuotaWarningLevel]`
: The warning level for the resource quota. Default `0`  (applies to type `ClassicSite`)

`--storageQuota [storageQuota]`
: The storage quota for this site collection in megabytes. Default `100`  (applies to type `ClassicSite`)

`--storageQuotaWarningLevel [storageQuotaWarningLevel]`
: The warning level for the storage quota in megabytes. Default `100`  (applies to type `ClassicSite`)

`--removeDeletedSite`
: Set, to remove existing deleted site with the same URL from the Recycle Bin  (applies to type `ClassicSite`)

`--wait`
: Wait for the site to be provisioned before completing the command  (applies to type `ClassicSite`)

`--disableFlows [disableFlows]`
: Set to `true` to disable using Microsoft Flow in this site collection

`--sharingCapability [sharingCapability]`
: The sharing capability for the site. Allowed values:  `Disabled`, `ExternalUserSharingOnly`, `ExternalUserAndGuestSharing`, `ExistingExternalUserSharingOnly`.

--8<-- "docs/cmd/_global.md"

## Remarks

This command checks if a site collection exists at the specified URL.

If no site is found, this command will create one using the specified options. For more information about creating site collections and things that you should into account, see the [`spo site add`](./site-add.md) command.

If a site is found at the specified URL, this command will update its properties using the specified values. For more information about updating site collections and things that you should take into account see the [`spo site set`](./site-set.md) command.

If another site exists at the specified URL and you haven't specified the `type` option, this command will proceed with updating its properties. If you specified the `type` and it doesn't match the type of the existing site, this command will return an error.

If you set the type to `ClassicSite`, you should also set a value of the `webTemplate` option. If you don't, checking the type of existing site will be skipped. If no site exists at the specified URL, creating the site will fail.

## Examples

Ensure that a site exists at the specified URL. Create a modern team site if no site exists. If a site is found, don't check its type and update its title.

```sh
m365 spo site ensure --url https://contoso.sharepoint.com/sites/team1 --alias team1 --title "Team 1"
```

Ensure that a communication site exists at the specified URL. Create a communication site if no site exists. If a site of different type is found, return an error.

```sh
m365 spo site ensure --url https://contoso.sharepoint.com/sites/comms --title Comms --type CommunicationSite
```

Ensure that a classic site exists at the specified URL. Create a classic team site if no site exists. If a site of different type is found, return an error.

```sh
m365 spo site ensure --url https://contoso.sharepoint.com/sites/classic --title Classic --type ClassicSite
```

Ensure that a site exists at the specified URL. Create a modern team site if no site exists. If a site is found, don't check its type and update its title and properties.

```sh
m365 spo site ensure --url https://contoso.sharepoint.com/sites/team1 --alias team1 --title "Team 1" --isPublic --shareByEmailEnabled
```

## More information

- [`spo site add`](./site-add.md)
- [`spo site set`](./site-set.md)
