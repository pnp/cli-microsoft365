# spo serviceprincipal set

Enable or disable the service principal

## Usage

```sh
m365 spo serviceprincipal set [options]
```

## Alias

```sh
m365 spo sp set
```

## Options

`-e, --enabled <enabled>`
: Set to `true` to enable the service principal or to `false` to disable it. Valid values are `true`, `false`.

`--confirm`
: Don't prompt for confirming enabling/disabling the service principal.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator.

Using the `-e, --enabled` option you can specify whether the service principal should be enabled or disabled. Use `true` to enable the service principal and `false` to disable it.

## Examples

Enable the service principal. Will prompt for confirmation

```sh
m365 spo serviceprincipal set --enabled true
```

Disable the service principal. Will prompt for confirmation

```sh
m365 spo serviceprincipal set --enabled false
```

Enable the service principal without prompting for confirmation

```sh
m365 spo serviceprincipal set --enabled true --confirm
```

## Response

=== "JSON"

    ```json
    {
      "AccountEnabled": true,
      "AppId": "48535560-3cc0-442e-a1b4-94c084b3ff59",
      "ReplyUrls": [
        "https://fluidpreview.office.net/spfxsinglesignon",
        "https://dev.fluidpreview.office.net/spfxsinglesignon",
        "https://contoso.sharepoint.com/",
        "https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx",
        "https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx?redirect",
        "https://contoso-admin.sharepoint.com/_forms/spfxsinglesignon.aspx"
      ]
    }
    ```

=== "Text"

    ```text
    AccountEnabled: true
    AppId         : 48535560-3cc0-442e-a1b4-94c084b3ff59
    ReplyUrls     : ["https://fluidpreview.office.net/spfxsinglesignon","https://dev.fluidpreview.office.net/spfxsinglesignon","https://contoso.sharepoint.com/","https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx","https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx?redirect","https://contoso-admin.sharepoint.com/_forms/spfxsinglesignon.aspx"]
    ```

=== "CSV"

    ```csv
    AccountEnabled,AppId,ReplyUrls
    1,48535560-3cc0-442e-a1b4-94c084b3ff59,"[""https://fluidpreview.office.net/spfxsinglesignon"",""https://dev.fluidpreview.office.net/spfxsinglesignon"",""https://contoso.sharepoint.com/"",""https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx"",""https://contoso.sharepoint.com/_forms/spfxsinglesignon.aspx?redirect"",""https://contoso-admin.sharepoint.com/_forms/spfxsinglesignon.aspx""]"
    ```
