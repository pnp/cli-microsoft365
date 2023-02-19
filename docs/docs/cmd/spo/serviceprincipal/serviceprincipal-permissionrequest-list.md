# spo serviceprincipal permissionrequest list

Lists pending permission requests

## Usage

```sh
m365 spo serviceprincipal permissionrequest list [options]
```

## Alias

```sh
m365 spo sp permissionrequest list
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    The admin role that's required to list permissions depends on the API. To approve permissions to any of the third-party APIs registered in the tenant, the application administrator role is sufficient. To approve permissions for Microsoft Graph or any other Microsoft API, the Global Administrator role is required.

## Examples

List all pending permission requests

```sh
m365 spo serviceprincipal permissionrequest list
```

## Response

=== "JSON"

    ```json
    [
      {
        "Id": "6eceed61-77e4-424d-ae1d-696a0de4d768",
        "Resource": "Microsoft Graph",
        "ResourceId": "Microsoft Graph",
        "Scope": "Reports.Read.All"
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Resource         ResourceId       Scope
    ------------------------------------  ---------------  ---------------  -----------------------
    6eceed61-77e4-424d-ae1d-696a0de4d768  Microsoft Graph  Microsoft Graph  Reports.Read.All
    ```

=== "CSV"

    ```csv
    Id,Resource,ResourceId,Scope
    6eceed61-77e4-424d-ae1d-696a0de4d768,Microsoft Graph,Microsoft Graph,Reports.Read.All
    ```
