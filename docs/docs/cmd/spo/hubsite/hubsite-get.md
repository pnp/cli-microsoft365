# spo hubsite get

Gets information about the specified hub site

## Usage

```sh
m365 spo hubsite get [options]
```

## Options

`-i, --id [id]`
: ID of the hubsite. Specify either `id`, `title` or `url` but not multiple.

`-t, --title [title]`
: Title of the hubsite. Specify either `id`, `title` or `url` but not multiple.

`-u, --url [url]`
: URL of the hubsite. Specify either `id`, `title` or `url` but not multiple.

`--includeAssociatedSites`
: Include the associated sites in the result (only in JSON output)

--8<-- "docs/cmd/_global.md"

## Examples

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_

```sh
m365 spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a
```

Get information about the hub site with Title _My Hub Site_

```sh
m365 spo hubsite get --title 'My Hub Site'
```

Get information about the hub site with URL _https://contoso.sharepoint.com/sites/HubSite_

```sh
m365 spo hubsite get --url 'https://contoso.sharepoint.com/sites/HubSite'
```

Get information about the hub site with ID _2c1ba4c4-cd9b-4417-832f-92a34bc34b2a_, including its associated sites. Associated site info is only shown in JSON output.

```sh
m365 spo hubsite get --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --includeAssociatedSites --output json
```

Get information about the hub site with Title _My Hub Site_

```sh
m365 spo hubsite get --title "My Hub Site"
```

Get information about the hub site with URL _https://contoso.sharepoint.com/sites/HubSite_

```sh
m365 spo hubsite get --url "https://contoso.sharepoint.com/sites/HubSite"
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)
