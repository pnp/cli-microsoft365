# spo propertybag get

Gets the value of the specified property from the property bag

## Usage

```sh
m365 spo propertybag get [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site from which the property bag value should be retrieved

`-k, --key <key>`
: Key of the property for which the value should be retrieved. Case-sensitive

`-f, --folder [folder]`
: Site-relative URL of the folder from which to retrieve property bag value. Case-sensitive

--8<-- "docs/cmd/_global.md"

## Examples

Returns the value of the _key1_ property from the property bag located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag get --webUrl https://contoso.sharepoint.com/sites/test --key key1
```

Returns the value of the _key1_ property from the property bag located in site root folder _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag get --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /
```

Returns the value of the _key1_ property from the property bag located in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag get --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents'
```

Returns the value of the _key1_ property from the property bag located in folder in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag get --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents/MyFolder'
```

Returns the value of the _key1_ property from the property bag located in site list _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag get --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /Lists/MyList
```
