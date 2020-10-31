# spo propertybag list

Gets property bag values

## Usage

```sh
m365 spo propertybag list [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: The URL of the site from which the property bag value should be retrieved

`-f, --folder [folder]`
: Site-relative URL of the folder from which to retrieve property bag value. Case-sensitive

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Return property bag values located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test
```

Return property bag values located in site root folder _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder /
```

Return property bag values located in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents'
```

Return property bag values located in folder in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder '/Shared Documents/MyFolder'
```

Return property bag values located in site list _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag list --webUrl https://contoso.sharepoint.com/sites/test --folder /Lists/MyList
```