# spo cdn origin remove

Removes CDN origin for the current SharePoint Online tenant

## Usage

```sh
m365 spo cdn origin remove [options]
```

## Options

`-h, --help`
: output usage information

`-t, --type [type]`
: Type of CDN to manage. `Public,Private`. Default `Public`

`-r, --origin <origin>`
: Origin to remove from the current CDN configuration

`--confirm`
: Don't prompt for confirming removal of a tenant property

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

Remove _*/CDN_ from the list of origins of the Public CDN

```sh
m365 spo cdn origin remove -t Public -r */CDN
```

## More information

- General availability of Microsoft 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
