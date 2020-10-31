# spo cdn policy set

Sets CDN policy value for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn policy set [options]
```

## Options

`-h, --help`
: output usage information

`-t, --type [type]`
: Type of CDN to manage. `Public,Private`. Default `Public`

`-p, --policy <policy>`
: CDN policy to configure. `IncludeFileExtensions|ExcludeRestrictedSiteClassifications`

`-v, --value <value>`
: Value for the policy to configure

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

Using the `-t, --type` option you can choose whether you want to manage the settings of the Public (default) or Private CDN. If you don't use the option, the command will use the Public CDN.

## Examples

Set the list of extensions supported by the Public CDN

```sh
m365 spo cdn policy set -t Public -p IncludeFileExtensions -v CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF,JSON
```

## More information

- General availability of Microsoft 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)