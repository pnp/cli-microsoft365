# spo site appcatalog remove

Removes a site collection app catalog in the specified site

## Usage

```sh
spo site appcatalog remove --url <url>
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site collection containing the app catalog to disable
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To remove an app catalog from a site collection, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

While the command uses the term *"remove"*, like the PowerShell equivalent cmdlet, it does not remove the special library **Apps for SharePoint** from the site collection. It simply disables the site collection app catalog in that site. Packages deployed to the app catalog are not available within the site collection.

## Examples

Remove app catalog for specified site collection.

```sh
spo site appcatalog remove --url https://contoso.sharepoint/sites/site
```

## More information

- Use the site collection app catalog: [https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog](https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog)