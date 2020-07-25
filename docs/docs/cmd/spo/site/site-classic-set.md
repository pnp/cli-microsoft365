# spo site classic set

Change classic site settings

## Usage

```sh
m365 spo site classic set [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: The absolute site url

`-t, --title [title]`
: The site title

`--sharing [sharing]`
: Sharing capabilities for the site. Allowed values: `Disabled,ExternalUserSharingOnly,ExternalUserAndGuestSharing,ExistingExternalUserSharingOnly`

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

`--owners [owners]`
: Comma-separated list of users to add as site collection administrators

`--lockState [lockState]`
: Sets site's lock state. Allowed values `Unlock,NoAdditions,ReadOnly,NoAccess`

`--noScriptSite [noScriptSite]`
: Specifies if the site allows custom script or not

`--wait`
: Wait for the settings to be applied before completing the command

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

The value of the `--resourceQuota` option must not exceed the company's aggregate available Sandboxed Solutions quota. For more information, see Resource Usage Limits on Sandboxed Solutions in SharePoint 2010: [http://msdn.microsoft.com/en-us/library/gg615462.aspx](http://msdn.microsoft.com/en-us/library/gg615462.aspx).

The value of the `--resourceQuotaWarningLevel` option must not exceed the value of the `--resourceQuota` option or the current value of the _UserCodeMaximumLevel_ property.

The value of the `--storageQuota` option must not exceed the company's available quota.

The value of the `--storageQuotaWarningLevel` option must not exceed the the value of the `--storageQuota` option or the current value of the _StorageMaximumLevel_ property.

When updating site owners using the `--owners` option, the command doesn't remove existing users but adds the users specified in the option to the list of already configured owners. When specifying owners, you can specify both users and groups.

For more information on locking classic sites see [https://technet.microsoft.com/en-us/library/cc263238.aspx](https://technet.microsoft.com/en-us/library/cc263238.aspx).

For more information on configuring no script sites see [https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f](https://support.office.com/en-us/article/Turn-scripting-capabilities-on-or-off-1f2c515f-5d7e-448a-9fd7-835da935584f).

Setting site properties is by default asynchronous and depending on the current state of Microsoft 365, might take up to few minutes. If you're building a script with steps that require the site to be fully configured, you should use the `--wait` flag. When using this flag, the `spo site classic set` command will keep running until it received confirmation from Microsoft 365 that the site has been fully configured.

## Examples

Change the title of the site collection. Don't wait for the configuration to complete

```sh
m365 spo site classic set --url https://contoso.sharepoint.com/sites/team --title Team
```

Add the specified user accounts as site collection administrators

```sh
m365 spo site classic set --url https://contoso.sharepoint.com/sites/team --owners "joe@contoso.com,steve@contoso.com"
```

Lock the site preventing users from accessing it. Wait for the configuration to complete

```sh
m365 spo site classic set --url https://contoso.sharepoint.com/sites/team --LockState NoAccess --wait
```
