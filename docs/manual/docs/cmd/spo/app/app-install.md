# spo app install

Installs an app from the tenant app catalog in the site

## Usage

```sh
spo app install [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the app to retrieve information for
`-s, --siteUrl <siteUrl>`|Absolute URL of the site to install the app in
`--verbose`|Runs command with verbose logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To install an app from the tenant app catalog in a site, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the app with the specified ID doesn't exist in the tenant app catalog, the command will fail with an error. Before you can install app in a site, you have to add it to the tenant app catalog first using the [spo app add](./app-add.md) command.

## Examples

```sh
spo app install -i b2307a39-e878-458b-bc90-03bc578531d6 -s https://contoso.sharepoint.com
```

Installs the app with ID _b2307a39-e878-458b-bc90-03bc578531d6_ in the _https://contoso.sharepoint.com_ site.

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)