# spo storageentity get

Get details for the specified tenant property

## Usage

```sh
spo storageentity get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-k, --key <key>`|Name of the tenant property to retrieve
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

Tenant properties are stored in the app catalog site associated with the site to which you are currently connected. When retrieving the specified tenant property, SharePoint will automatically find the associated app catalog and try to retrieve the property from it.

## Examples

Show the value, description and comment of the _AnalyticsId_ tenant property

```sh
spo storageentity get -k AnalyticsId
```

## More information

- SharePoint Framework Tenant Properties: [https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/tenant-properties)
