# spo user Add

Adds a site user within specific web

## Usage

```sh
spo user add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl<webUrl>` |Url of the web to list the users within
`--group [group]`|The SharePoint group name to which the user to be added
`--email <email>`|Email of the user to be added
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Adds user with email _john.doe@mytenant.onmicrosoft.com_ to web _https://contoso.sharepoint.com/sites/HR_

```sh
spo user add --webUrl "https://contoso.sharepoint.com/sites/mysite" --email "john.doe@contoso.onmicrosoft.com" --group "HR Members"
```

## More information

- Add-PnPUserToGroup - [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpusertogroup?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/add-pnpusertogroup?view=sharepoint-ps)


