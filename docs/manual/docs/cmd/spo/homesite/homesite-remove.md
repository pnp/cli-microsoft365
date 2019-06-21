# spo homesite remove

Removes the the Home Site

## Usage

```sh
spo homesite remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`--confirm`|Do not prompt for confirmation before deleting the Home Site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To remove the Home Site, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Removes the current Home Site without confirmation

```sh
spo homesite remove --confirm
```

## More information

- SharePoint home sites: a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)
