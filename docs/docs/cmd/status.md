# status

Shows Microsoft 365 login status

## Usage

```sh
m365 status [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

If you are logged in to Microsoft 365, the `status` command will show you information about the user or application name used to sign in and the details about the stored refresh and access tokens and their expiration date and time when run in debug mode.

## Examples

Show the information about the current login to the Microsoft 365

```sh
m365 status
```

## Response

=== "JSON"

    ```json
    {
      "connectedAs": "john.doe@contoso.onmicrosoft.com",
      "authType": "DeviceCode",
      "appId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
      "appTenant": "common"
    }
    ```

=== "Text"

    ```text
    appId      : 31359c7f-bd7e-475c-86db-fdb8c937548e
    appTenant  : common
    authType   : DeviceCode
    connectedAs: john.doe@contoso.onmicrosoft.com
    ```

=== "CSV"

    ```csv
    connectedAs,authType,appId,appTenant
    john.doe@contoso.onmicrosoft.com,DeviceCode,31359c7f-bd7e-475c-86db-fdb8c937548e,common
    ```

=== "Markdown"

    ```md
    # status

    Date: 7/2/2023



    Property | Value
    ---------|-------
    connectedAs | john.doe@contoso.onmicrosoft.com
    authType | DeviceCode
    appId | 31359c7f-bd7e-475c-86db-fdb8c937548e
    appTenant | common
    ```
