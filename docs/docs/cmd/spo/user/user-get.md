# spo user get

Gets a site user within specific web

## Usage

```sh
m365 spo user get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the web to get the user within

`-i, --id [id]`
: ID of the user to retrieve information for. Use either `email`, `id` or `loginName`, but not all.

`--email [email]`
: Email of the user to retrieve information for. Use either `email`, `id` or `loginName`, but not all.

`--loginName [loginName]`
: Login name of the user to retrieve information for. Use either `email`, `id` or `loginName`, but not all.

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get user with email _john.doe@mytenant.onmicrosoft.com_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --email john.doe@mytenant.onmicrosoft.com
```

Get user with ID _6_ for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --id 6
```

Get user with login name 'i:0#.f
: membership|john.doe@mytenant.onmicrosoft.com' for web _https://contoso.sharepoint.com/sites/project-x_

```sh
m365 spo user get --webUrl https://contoso.sharepoint.com/sites/project-x --loginName "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
```
