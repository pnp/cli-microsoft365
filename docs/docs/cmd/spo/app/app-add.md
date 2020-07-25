# spo app add

Adds an app to the specified SharePoint Online app catalog

## Usage

```sh
m365 spo app add [options]
```

## Options

`-h, --help`
: output usage information

`-p, --filePath <filePath>`
: Absolute or relative path to the solution package file to add to the app catalog

`--overwrite`
: Set to overwrite the existing package file

`-s, --scope [scope]`
: Scope of the app catalog: `tenant,sitecollection`. Default `tenant`

`-u, --appCatalogUrl [appCatalogUrl]`
: The URL of the app catalog where the solution package will be added. It must be specified when the scope is `sitecollection`

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

When specifying the path to the app package file you can use both relative and absolute paths. Note, that `~` in the path, will not be resolved and will most likely result in an error.

When adding an app to the tenant app catalog, it's not necessary to specify the tenant app catalog URL. When the URL is not specified, the CLI will try to resolve the URL itself. Specifying the app catalog URL is required when you want to add the app to a site collection app catalog.

When specifying site collection app catalog, you can specify the URL either with our without the _AppCatalog_ part, for example `https://contoso.sharepoint.com/sites/team-a/AppCatalog` or `https://contoso.sharepoint.com/sites/team-a`. CLI will accept both formats.

If you try to upload a package that already exists in the app catalog without specifying the `--overwrite` option, the command will fail with an error stating that the specified package already exists.

## Examples

Add the _spfx.sppkg_ package to the tenant app catalog

```sh
m365 spo app add --filePath /Users/pnp/spfx/sharepoint/solution/spfx.sppkg
```

Overwrite the _spfx.sppkg_ package in the tenant app catalog with the newer version

```sh
m365 spo app add --filePath sharepoint/solution/spfx.sppkg --overwrite
```

Add the _spfx.sppkg_ package to the site collection app catalog of site _https://contoso.sharepoint.com/sites/site1_

```sh
m365 spo app add --filePath c:\spfx.sppkg --scope sitecollection --appCatalogUrl https://contoso.sharepoint.com/sites/site1
```

## More information

- Application Lifecycle Management (ALM) APIs: [https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins](https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins)