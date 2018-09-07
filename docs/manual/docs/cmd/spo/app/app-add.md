# spo app add

Adds an app to the specified SharePoint Online app catalog

## Usage

```sh
spo app add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --filePath <filePath>`|Absolute or relative path to the solution package file to add to the app catalog
`--overwrite`|Set to overwrite the existing package file
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To add an app to the tenant app catalog, you have to first log in to a SharePoint site using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

When specifying the path to the app package file you can use both relative and absolute paths. Note, that `~` in the path, will not be resolved and will most likely result in an error.

If you try to upload a package that already exists in the tenant app catalog without specifying the `--overwrite` option, the command will fail with an error stating that the specified package already exists.

## Examples

Add the _spfx.sppkg_ package to the tenant app catalog

```sh
spo app add -p /Users/pnp/spfx/sharepoint/solution/spfx.sppkg
```

Overwrite the _spfx.sppkg_ package in the tenant app catalog with the newer version

```sh
spo app add -p sharepoint/solution/spfx.sppkg --overwrite
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)