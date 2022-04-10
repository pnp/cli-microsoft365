# spo propertybag remove

Removes specified property from the property bag

## Usage

```sh
m365 spo propertybag remove [options]
```

## Options

`-u, --webUrl <webUrl>`
: The URL of the site from which the property should be removed

`-k, --key <key>`
: Key of the property to be removed. Case-sensitive

`-f, --folder [folder]`
: Site-relative URL of the folder from which to remove the property bag value

`--confirm`
: Don't prompt for confirming removal of property bag value

--8<-- "docs/cmd/_global.md"

## Examples

Removes the value of the _key1_ property from the property bag located in site _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag remove --webUrl https://contoso.sharepoint.com/sites/test --key key1
```

Removes the value of the _key1_ property from the property bag located in site root folder _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag remove --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder / --confirm
```

Removes the value of the _key1_ property from the property bag located in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag remove --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents'
```

Removes the value of the _key1_ property from the property bag located in folder in site document library _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag remove --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder '/Shared Documents/MyFolder'
```

Removes the value of the _key1_ property from the property bag located in site list _https://contoso.sharepoint.com/sites/test_

```sh
m365 spo propertybag remove --webUrl https://contoso.sharepoint.com/sites/test --key key1 --folder /Lists/MyList
```
