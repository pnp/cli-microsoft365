# spo customaction clear

Deletes all custom actions from site or site collection

## Usage

```sh
m365 spo customaction clear [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: Url of the site or site collection to clear the custom actions from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

`--confirm`
: Don't prompt for confirming removing all custom actions

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Clears all user custom actions for both site and site collection _https://contoso.sharepoint.com/sites/test_.
Skips the confirmation prompt message.

```sh
m365 spo customaction clear -u https://contoso.sharepoint.com/sites/test --confirm
```

Clears all user custom actions for site _https://contoso.sharepoint.com/sites/test_. 

```sh
m365 spo customaction clear -u https://contoso.sharepoint.com/sites/test -s Web
```

Clears all user custom actions for site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction clear --url https://contoso.sharepoint.com/sites/test --scope Site
```