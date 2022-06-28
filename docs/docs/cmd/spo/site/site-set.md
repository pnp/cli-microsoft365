# spo site set

Updates properties of the specified site

## Usage

```sh
m365 spo site set [options]
```

## Options

`-u, --url <url>`
: The URL of the site collection to update

`-t, --title [title]`
: The new title for the site collection

`-d, --description [description]`
: The site description

`--classification [classification]`
: The new classification for the site collection

`--disableFlows [disableFlows]`
: Set to `true` to disable using Microsoft Flow in this site collection

`--isPublic [isPublic]`
: Set to `true` to make the group linked to the site public or to `false` to make it private

`--owners [owners]`
: Comma-separated list of users to add as site collection administrators

`--shareByEmailEnabled [shareByEmailEnabled]`
: Set to true to allow to share files with guests and to false to disallow it

`--siteDesignId [siteDesignId]`
: Id of the custom site design to apply to the site

`--sharingCapability [sharingCapability]`
: The sharing capability for the site. Allowed values:  `Disabled`, `ExternalUserSharingOnly`, `ExternalUserAndGuestSharing`, `ExistingExternalUserSharingOnly`.

`--siteLogoUrl [siteLogoUrl]`
: Set the logo for the site collection. This can be an absolute or relative URL to a file on the current site collection.

`--resourceQuota [resourceQuota]`
: The quota for this site collection in Sandboxed Solutions units

`--resourceQuotaWarningLevel [resourceQuotaWarningLevel]`
: The warning level for the resource quota

`--storageQuota [storageQuota]`
: The storage quota for this site collection in megabytes

`--storageQuotaWarningLevel [storageQuotaWarningLevel]`
: The warning level for the storage quota in megabytes

`--allowSelfServiceUpgrade [allowSelfServiceUpgrade]`
: Set to allow tenant administrators to upgrade the site collection

`--lockState [lockState]`
: Sets site's lock state. Allowed values `Unlock,NoAdditions,ReadOnly,NoAccess`

`--noScriptSite [noScriptSite]`
: Specifies if the site allows custom script or not

`--wait`
: Wait for the settings to be applied before completing the command

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

If the specified url doesn't refer to an existing site collection, you will get a `404 - "404 FILE NOT FOUND"` error.

The `isPublic` property can be set only on groupified site collections. If you try to set it on a site collection without a group, you will get an error.

When setting owners, the specified owners will be added to the already configured owners. Existing owners will not be removed.

The value of the `--resourceQuota` option must not exceed the company's aggregate available Sandboxed Solutions quota. For more information, see Resource Usage Limits on Sandboxed Solutions in SharePoint 2010: [http://msdn.microsoft.com/en-us/library/gg615462.aspx](http://msdn.microsoft.com/en-us/library/gg615462.aspx).

The value of the `--resourceQuotaWarningLevel` option must not exceed the value of the `--resourceQuota` option or the current value of the _UserCodeMaximumLevel_ property.

The value of the `--storageQuota` option must not exceed the company's available quota.

The value of the `--storageQuotaWarningLevel` option must not exceed the the value of the `--storageQuota` option or the current value of the _StorageMaximumLevel_ property.

For more information on locking sites see [https://technet.microsoft.com/en-us/library/cc263238.aspx](https://technet.microsoft.com/en-us/library/cc263238.aspx).

For more information on configuring no script sites see [https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f](https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f).

Setting site properties is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully configured, you should use the `--wait` flag. When using this flag, the `spo site set` command will keep running until it received confirmation from Microsoft 365 that the site has been fully configured.

## Examples

Update site collection's classification

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --classification MBI
```

Reset site collection's classification.

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --classification
```

Disable using Microsoft Flow on the site collection

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --disableFlows true
```

Update the visibility of the Microsoft 365 group behind the specified groupified site collection to public

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --isPublic true
```

Update site collection's owners

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --owners "john@contoso.onmicrosoft.com,steve@contoso.onmicrosoft.com"
```

Allow sharing files in the site collection with guests

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --shareByEmailEnabled true
```

Apply the specified site ID to the site collection

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --siteDesignId "eb2f31da-9461-4fbf-9ea1-9959b134b89e"
```

Update site collection's title

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --title "My new site"
```

Update site collection's description

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --description "my description"
```

Restrict external sharing to already available external users only

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/sales --sharingCapability ExistingExternalUserSharingOnly
```

Set the logo on the site

```sh
m365 spo site set --url https://estruyfdev2.sharepoint.com/sites/sales --siteLogoUrl "/sites/sales/SiteAssets/parker-ms-1200.png"
```

Unset the logo on the site

```sh
m365 spo site set --url https://estruyfdev2.sharepoint.com/sites/sales --siteLogoUrl ""
```

Lock the site preventing users from accessing it. Wait for the configuration to complete

```sh
m365 spo site set --url https://contoso.sharepoint.com/sites/team --LockState NoAccess --wait
```
