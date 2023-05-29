# app permission list

Lists API permissions for the current AAD app

## Usage

```sh
m365 app permission list [options]
```

## Options

`--appId [appId]`
: Client ID of the Azure AD app registered in the .m365rc.json file to retrieve API permissions for.

--8<-- "docs/cmd/_global.md"

## Remarks

Use this command to quickly look up API permissions for the Azure AD application registration registered in the _.m365rc.json_ file in your current project (folder).

If you have multiple apps registered in your .m365rc.json file, you can specify the app for which you'd like to retrieve permissions using the `--appId` option. If you don't specify the app using the `--appId` option, you'll be prompted to select one of the applications from your _.m365rc.json_ file.

## Examples

Retrieve API permissions for your current Azure AD app.

```sh
m365 app permission list
```

Retrieve API permissions for the Azure AD app with the client ID specified in the _.m365rc.json_ file.

```sh
m365 app permission list --appId e23d235c-fcdf-45d1-ac5f-24ab2ee0695d
```

## Response

=== "JSON"

    ```json
    [
      {
        "resource": "Microsoft Teams - Teams And Channels Service",
        "permission": "channels.readwrite",
        "type": "Application"
      },
      {
        "resource": "Yammer",
        "permission": "access_as_user",
        "type": "Delegated"
      },
      {
        "resource": "Yammer",
        "permission": "user_impersonation",
        "type": "Delegated"
      }
    ]
    ```

=== "Text"

    ```text
    resource                                      permission          type
    --------------------------------------------  ------------------  -----------
    Microsoft Teams - Teams And Channels Service  channels.readwrite  Application
    Yammer                                        access_as_user      Delegated
    Yammer                                        user_impersonation  Delegated
    ```

=== "CSV"

    ```csv
    resource,permission,type
    Microsoft Teams - Teams And Channels Service,channels.readwrite,Application
    Yammer,access_as_user,Delegated
    Yammer,user_impersonation,Delegated
    ```

=== "Markdown"

    ```md
    # app permission list

    Date: 5/29/2023

    Property | Value
    ---------|-------
    resource | Microsoft Teams - Teams And Channels Service
    permission | channels.readwrite
    type | Application

    Property | Value
    ---------|-------
    resource | Yammer
    permission | access\_as\_user
    type | Delegated
    
    Property | Value
    ---------|-------
    resource | Yammer
    permission | user\_impersonation
    type | Delegated
    ```
