# spo serviceprincipal grant list

Lists permissions granted to the service principal

## Usage

```sh
m365 spo serviceprincipal grant list [options]
```

## Alias

```sh
m365 spo sp grant list
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you must be a Global administrator.

## Examples

List all permissions granted to the service principal

```sh
m365 spo serviceprincipal grant list
```

## Response

=== "JSON"

    ```json
    [
      {
        "IsDomainIsolated": false,
        "ObjectId": "QqYEYFwYmkeZKhXRwj4iKV5QwbD60RVCo6xeMUG407E",
        "PackageName": null,
        "Resource": "Windows Azure Active Directory",
        "ResourceId": "b0c1505e-d1fa-4215-a3ac-5e3141b8d3b1",
        "Scope": "User.Read"
      }
    ]
    ```

=== "Text"

    ```text
    IsDomainIsolated  ObjectId                                     PackageName  Resource                        ResourceId                            Scope
    ----------------  -------------------------------------------  -----------  ------------------------------  ------------------------------------  --------------------------
    false             QqYEYFwYmkeZKhXRwj4iKV5QwbD60RVCo6xeMUG407E  null         Windows Azure Active Directory  b0c1505e-d1fa-4215-a3ac-5e3141b8d3b1  User.Read
    ```

=== "CSV"

    ```csv
    IsDomainIsolated,ObjectId,PackageName,Resource,ResourceId,Scope
    ,QqYEYFwYmkeZKhXRwj4iKV5QwbD60RVCo6xeMUG407E,,Windows Azure Active Directory,b0c1505e-d1fa-4215-a3ac-5e3141b8d3b1,User.Read
    ```

=== "Markdown"

    ```md
    # spo serviceprincipal grant list 

    Date: 5/7/2023

    ## 4WtBzD8u5kW-sYuikIWL_8ZYTP5mJB1LnC6OT4Ibr94

    Property | Value
    ---------|-------
    IsDomainIsolated | false
    ObjectId | 4WtBzD8u5kW-sYuikIWL\_8ZYTP5mJB1LnC6OT4Ibr94
    Resource | Microsoft Graph
    ResourceId | fe4c58c6-2466-4b1d-9c2e-8e4f821bafde
    Scope | Mail.Read
    ```
