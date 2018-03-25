# spo propertybag set

Sets the value of the specified property in the property bag. Adds the property if it does not exist

## Usage

```sh
spo propertybag set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|The URL of the site in which the property should be set
`-k, --key <key>`|Key of the property to be set. Case-sensitive
`-v, --value <value>`|Value of the property to be set
`-f, --folder [folder]`|Site-relative URL of the folder on which the property should be set
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To set property bag value, you have to first connect to a SharePoint Online site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

## Examples

Sets the value of the _key1_ property in the property bag of site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1
```

Sets the value of the _key1_ property in the property bag of the root folder of site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /
```

Sets the value of the _key1_ property in the property bag of a document library located in site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents'
```

Sets the value of the _key1_ property in the property bag of a folder in a document library located in site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents/MyFolder'
```

Sets the value of the _key1_ property in the property bag of a list in site _https://contoso.sharepoint.com/sites/test_

```sh
spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /Lists/MyList
```