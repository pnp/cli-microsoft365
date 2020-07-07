# spo group get

Gets a site group within specific web

## Usage

```sh
spo group get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the web to get the group within
`--id [id]`|ID of the site group to retrieve information for. Use either "id" or "name", but not all. e.g '7'
`--name [name]`|Name of the site group to retrieve information for. Specify either `id` or `name` but not both e.g 'Team Site Members'
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get group with ID 7 for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --id 7 

Get group with name _Team Site Members_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo group get --webUrl https://contoso.sharepoint.com/sites/project-x --name "Team Site Members"
```


## More information

- Get-PnPGroup - [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpgroup?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpgroup?view=sharepoint-ps)