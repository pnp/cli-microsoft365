# spo cdn set

Enable or disable the specified Microsoft 365 CDN

## Usage

```sh
m365 spo cdn set [options]
```

## Options

`-h, --help`
: output usage information

`-e, --enabled <enabled>`
: Set to true to enable CDN or to false to disable it. Valid values are `true,false`

`-t, --type [type]`
: Type of CDN to manage. `Public,Private,Both`. Default `Public`

`--noDefaultOrigins`
: pass to not create the default Origins

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

Using the `-t, --type` option you can choose whether you want to manage the settings of the Public (default), Private CDN or both. If you don't use the option, the command will use the Public CDN.

Using the `-e, --enabled` option you can specify whether the given CDN type should be enabled or disabled. Use true to enable the specified CDN and false to disable it.

Using the `--noDefaultOrigins` option you can specify to skip the creation of the default origins.

## Examples

Enable the Microsoft 365 Public CDN on the current tenant

```sh
m365 spo cdn set -t Public -e true
```

Disable the Microsoft 365 Public CDN on the current tenant

```sh
m365 spo cdn set -t Public -e false
```

Enable the Microsoft 365 Private CDN on the current tenant

```sh
m365 spo cdn set -t Private -e true
```

Enable the Microsoft 365 Private and Public CDN on the current tenant with default origins

```sh
m365 spo cdn set -t Both -e true
```

Enable the Microsoft 365 Private and Public CDN on the current tenant without default origins

```sh
m365 spo cdn set -t Both -e true --noDefaultOrigins
```

## More information

- General availability of Microsoft 365 CDN: [https://dev.office.com/blogs/general-availability-of-office-365-cdn](https://dev.office.com/blogs/general-availability-of-office-365-cdn)
