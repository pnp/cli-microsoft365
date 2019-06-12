# spo homesite set

Sets the specified site as the Home Site

## Usage

```sh
spo homesite set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --siteUrl <siteUrl>`|The URL of the site to set as Home Site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To set site as the Home Site, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Set the specified site as the Home Site

```sh
spo homesite set --siteUrl https://contoso.sharepoint.com/sites/comms
```

## More information

- SharePoint home sites: a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)