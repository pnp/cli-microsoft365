# spo user remove

Removes user from specific web

## Usage

```sh
m365 spo user remove [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the web to remove user

`--id [id]`
: ID of the user to remove from web

`--loginName [loginName]`
: Login name of the site user to remove

`--confirm`
: Do not prompt for confirmation before removing user from web

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

Use either `id` or `loginName`, but not both

## Examples

Removes user with id _10_ from web _https://contoso.sharepoint.com/sites/HR_ without prompting for confirmation

```sh
m365 spo user remove --webUrl "https://contoso.sharepoint.com/sites/HR" --id 10 --confirm
```

Removes user with login name _i:0#.f|membership|john.doe@mytenant.onmicrosoft.com_ from web _https://contoso.sharepoint.com/sites/HR_

```sh
m365 spo user remove --webUrl "https://contoso.sharepoint.com/sites/HR" --loginName "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
```

## More information

- Remove-PnPUser - [https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/remove-pnpuser?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/remove-pnpuser?view=sharepoint-ps)