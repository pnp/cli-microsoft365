# spo customaction remove

Removes specified custom action from site or site collection

## Usage

```sh
m365 spo customaction remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --id <id>`
: Id (GUID) of the custom action to remove

`-u, --url <url>`
: Url of the site or site collection to remove the custom action from

`-s, --scope [scope]`
: Scope of the custom action. Allowed values `Site,Web,All`. Default `All`

`--confirm`
: Don't prompt for confirming removal of a user custom action

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site or site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site or site collection _https://contoso.sharepoint.com/sites/test_. Skips the confirmation prompt message.

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --confirm
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site collection _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove -i 058140e3-0e37-44fc-a1d3-79c487d371a3 -u https://contoso.sharepoint.com/sites/test -s Site
```

Removes user custom action with ID _058140e3-0e37-44fc-a1d3-79c487d371a3_ located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo customaction remove --id 058140e3-0e37-44fc-a1d3-79c487d371a3 --url https://contoso.sharepoint.com/sites/test --scope Web
```

## More information

- UserCustomAction REST API resources: [https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction](https://msdn.microsoft.com/en-us/library/office/dn531432.aspx#bk_UserCustomAction)