# spo user get

Gets a site user within specific web

## Usage

```sh
spo user get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|URL of the web to get the user within 
`--email [email]`|Email of the user to retrieve information for. Use either "email", "id" or "loginName", but not all. e.g 'john.doe@mytenant.onmicrosoft.com'
`--id [id]`|ID of the user to retrieve information for. Use either "email", "id" or "loginName", but not all. e.g '6'
`--loginName [loginName]`|loginName of the user to retrieve information for. Specify either `id` or `title` but not both e.g 'i:0#.f|membership|john.doe@mytenant.onmicrosoft.com'
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get user with email _john.doe@mytenant.onmicrosoft.com_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --email john.doe@mytenant.onmicrosoft.com
```

Get user with ID 6 for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --id 6 
```
Get user with LoginName 'i:0#.f|membership|john.doe@mytenant.onmicrosoft.com' for web _https://contoso.sharepoint.com/sites/project-x_

```sh
spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --loginName i:0#.f|membership|john.doe@mytenant.onmicrosoft.com 
```



## More information

- Get-PnPUser - [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpuser?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/get-pnpuser?view=sharepoint-ps)