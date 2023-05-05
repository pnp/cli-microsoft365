# cli doctor

Retrieves diagnostic information about the current environment

## Usage

```sh
m365 cli doctor [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

This command gets all the necessary diagnostic information needed to triage and debug CLI issues without exposing any security-sensitive details

## Examples

Retrieve diagnostic information

```sh
m365 cli doctor
```

## Response

=== "JSON"

    ```json
    {
      "os": {
        "platform": "win32",
        "version": "Windows 10 Pro",
        "release": "10.0.19045"
      },
      "cliVersion": "6.1.0",
      "nodeVersion": "v16.13.0",
      "cliAadAppId": "31359c7f-bd7e-475c-86db-fdb8c937548e",
      "cliAadAppTenant": "common",
      "authMode": "DeviceCode",
      "cliEnvironment": "",
      "cliConfig": {
        "output": "json",
        "showHelpOnFailure": false
      },
      "roles": [],
      "scopes": [
        "AllSites.FullControl"
      ]
    }
    ```

=== "Text"

    ```text
    authMode       : DeviceCode
    cliAadAppId    : 31359c7f-bd7e-475c-86db-fdb8c937548e
    cliAadAppTenant: common
    cliConfig      : {"output":"json","showHelpOnFailure":false}
    cliEnvironment :
    cliVersion     : 6.1.0
    nodeVersion    : v16.13.0
    os             : {"platform":"win32","version":"Windows 10 Pro","release":"10.0.19045"}
    roles          : []
    scopes         : ["AllSites.FullControl"]
    ```

=== "CSV"

    ```csv
    os,cliVersion,nodeVersion,cliAadAppId,cliAadAppTenant,authMode,cliEnvironment,cliConfig,roles,scopes
    "{""platform"":""win32"",""version"":""Windows 10 Pro"",""release"":""10.0.19045""}",6.1.0,v16.13.0,31359c7f-bd7e-475c-86db-fdb8c937548e,common,DeviceCode,,"{""output"":""json"",""showHelpOnFailure"":false}",[],"[""AllSites.FullControl""]"
    ```

=== "Markdown"

    ```md
    # cli doctor

    Date: 2022-09-05

    Property | Value
    ---------|-------
    cliVersion | 6.1.0
    nodeVersion | v16.13.0
    cliAadAppId | 31359c7f-bd7e-475c-86db-fdb8c937548e
    cliAadAppTenant | common
    authMode | DeviceCode
    cliEnvironment |
    ```
