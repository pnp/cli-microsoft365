# spo web add

Create new subsite

## Usage

```sh
m365 spo web add [options]
```

## Options

`-t, --title <title>`
: Subsite title

`-d, --description [description]`
: Subsite description

`-u, --url <url>`
: Subsite relative url

`-w, --webTemplate <webTemplate>`
: Subsite template, eg. `STS#0` (Classic team site)

`-p, --parentWebUrl <parentWebUrl>`
: URL of the parent site under which to create the subsite

`-l, --locale [locale]`
: Subsite locale LCID, eg. `1033` for en-US. Default `1033`

`--breakInheritance`
: Set to not inherit permissions from the parent site

`--inheritNavigation`
: Set to inherit the navigation from the parent site

--8<-- "docs/cmd/_global.md"

## Examples

Create subsite using the _Team site_ template in the _en-US_ locale

```sh
m365 spo web add --title Subsite --description Subsite --url subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --locale 1033
```

Create subsite with unique permissions using the default _en-US_ locale

```sh
m365 spo web add --title Subsite --url subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --breakInheritance
```

Create subsite with the same navigation as the parent site

```sh
m365 spo web add --title Subsite --url subsite --webTemplate STS#0 --parentWebUrl https://contoso.sharepoint.com --inheritNavigation
```

## Response

=== "JSON"

    ```json
    {
      "Configuration": 0,
      "Created": "2022-11-05T14:07:51",
      "Description": "Subsite",
      "Id": "b60137df-c3dc-4984-9def-8edcf7c98ab9",
      "Language": 1033,
      "LastItemModifiedDate": "2022-11-05T14:08:03Z",
      "LastItemUserModifiedDate": "2022-11-05T14:08:03Z",
      "ServerRelativeUrl": "/subsite",
      "Title": "Subsite",
      "WebTemplate": "STS",
      "WebTemplateId": 0
    }
    ```

=== "Text"

    ```text
    Configuration           : 0
    Created                 : 2022-11-05T14:08:35
    Description             : Subsite
    Id                      : 1f2db405-394d-470e-b820-cd5182883b45
    Language                : 1033
    LastItemModifiedDate    : 2022-11-05T14:08:47Z
    LastItemUserModifiedDate: 2022-11-05T14:08:47Z
    ServerRelativeUrl       : /subsite
    Title                   : Subsite
    WebTemplate             : STS
    WebTemplateId           : 0
    ```

=== "CSV"

    ```csv
    Configuration,Created,Description,Id,Language,LastItemModifiedDate,LastItemUserModifiedDate,ServerRelativeUrl,Title,WebTemplate,WebTemplateId
    0,2022-11-05T14:09:11,Subsite,0cbf2896-bac2-4244-b871-68b413ee7b2f,1033,2022-11-05T14:09:22Z,2022-11-05T14:09:22Z,/subsite,Subsite,STS,0
    ```

=== "Markdown"

    ```md
    # spo web add --title "Subsite" --url "subsite" --webTemplate "STS#0" --parentWebUrl "https://contoso.sharepoint.com" --inheritNavigation "true"

    Date: 4/10/2023

    ## Subsite (261ab6d3-0064-47d8-9189-82e5745d7a7f)

    Property | Value
    ---------|-------
    Configuration | 0
    Created | 2023-04-10T06:38:24
    Description | 
    Id | 261ab6d3-0064-47d8-9189-82e5745d7a7f
    Language | 1033
    LastItemModifiedDate | 2023-04-10T06:38:32Z
    LastItemUserModifiedDate | 2023-04-10T06:38:32Z
    ServerRelativeUrl | /subsite
    Title | Subsite
    WebTemplate | STS
    WebTemplateId | 0
    ```
