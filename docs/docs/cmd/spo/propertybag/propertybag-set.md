# spo propertybag set

Sets the value of the specified property in the property bag. Adds the property if it does not exist

## Usage

```sh
m365 spo propertybag set [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site in which the property should be set

`-k, --key <key>`
: Key of the property to be set. Case-sensitive

`-v, --value <value>`
: Value of the property to be set

`-f, --folder [folder]`
: Site-relative URL of the folder on which the property should be set

--8<-- "docs/cmd/_global.md"

## Remarks

SharePoint Online supports setting property bag values only in classic sites. On modern sites you will get a _Site has NoScript enabled, and setting property bag values is not supported_ error.

## Examples

Sets the value of the property in the property bag of the given site

```sh
m365 spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1
```

Sets the value of the property in the property bag of the root folder of the given site

```sh
m365 spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /
```

Sets the value of the property in the property bag of a document library located in the given site

```sh
m365 spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents'
```

Sets the value of the property in the property bag of a folder in a document library located in the given site

```sh
m365 spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder '/Shared Documents/MyFolder'
```

Sets the value of the property in the property bag of a list in the given site

```sh
m365 spo propertybag set --webUrl https://contoso.sharepoint.com/sites/test --key key1 --value value1 --folder /Lists/MyList
```

## Response

The command won't return a response on success.
