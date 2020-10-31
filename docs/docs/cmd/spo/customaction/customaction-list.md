# spo customaction list

Lists user custom actions for site or site collection

## Usage

```sh
m365 spo customaction list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: Url of the site or site collection to retrieve the custom action from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When using the text output type (default), the command lists only the values of the `Name`, `Location`, `Scope` and `Id` properties of the custom action. When setting the output type to JSON, all available properties are included in the command output.

## Examples

Return details about all user custom actions located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction list -u https://contoso.sharepoint.com/sites/test
```

Return details about all user custom actions located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction list --url https://contoso.sharepoint.com/sites/test
```

Return details about all user custom actions located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction list -u https://contoso.sharepoint.com/sites/test -s Site
```

Return details about all user custom actions located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction list --url https://contoso.sharepoint.com/sites/test --scope Web
```

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)