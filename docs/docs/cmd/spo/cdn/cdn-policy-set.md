# spo cdn policy set

Sets CDN policy value for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn policy set [options]
```

## Options

`-t, --cdnType [cdnType]`
: Type of CDN to manage. `Public,Private`. Default `Public`

`-p, --policy <policy>`
: CDN policy to configure. `IncludeFileExtensions|ExcludeRestrictedSiteClassifications`

`-v, --value <value>`
: Value for the policy to configure

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --cdnType` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Set the list of extensions supported by the Public CDN

```sh
m365 spo cdn policy set --cdnType Public --policy IncludeFileExtensions --value CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON
```

## More information

- Use Microsoft 365 CDN with SharePoint Online: [https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide](https://learn.microsoft.com/microsoft-365/enterprise/use-microsoft-365-cdn-with-spo?view=o365-worldwide)
