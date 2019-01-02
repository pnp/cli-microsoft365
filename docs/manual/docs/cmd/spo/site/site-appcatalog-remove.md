# spo site appcatalog remove

Removes site collection app catalog from the specified site

## Usage

```sh
spo site appcatalog remove [options]
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
    Before using this command, log in to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To remove a site collection catalog, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

While the command uses the term *'remove'*, like its equivalent PowerShell cmdlet, it does not remove the special library **Apps for SharePoint** from the site collection. Instead, it disables the site collection app catalog in that site. Packages deployed to the app catalog are not available within the site collection.

## Examples

Remove the site collection app catalog from specified site

```sh
spo site appcatalog remove --url https://contoso.sharepoint/sites/site
```

## More information

- Use the site collection app catalog: [https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog](https://docs.microsoft.com/en-us/sharepoint/dev/general-development/site-collection-app-catalog)