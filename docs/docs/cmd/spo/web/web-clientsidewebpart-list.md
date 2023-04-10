# spo web clientsidewebpart list

Lists available client-side web parts

## Usage

```sh
m365 spo web clientsidewebpart list [options]
```

## Options

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve the information

--8<-- "docs/cmd/_global.md"

## Examples

Lists all the available client-side web parts for the specified site

```sh
m365 spo web clientsidewebpart list --webUrl https://contoso.sharepoint.com
```

## Response

=== "JSON"

    ```json
    [ 
      {
        "Id": "9cc0f495-db64-4d74-b06b-a3de16231fe1",
        "Name": "9cc0f495-db64-4d74-b06b-a3de16231fe1",
        "Title": "Dashboard for Viva Connections"
      }
    ]
    ```

=== "Text"

    ```text
    Id                                    Name                                  Title
    ------------------------------------  ------------------------------------  ------------------------------
    9cc0f495-db64-4d74-b06b-a3de16231fe1  9cc0f495-db64-4d74-b06b-a3de16231fe1  Dashboard for Viva Connections
    ```

=== "CSV"

    ```csv
    Id,Name,Title
    9cc0f495-db64-4d74-b06b-a3de16231fe1,9cc0f495-db64-4d74-b06b-a3de16231fe1,Dashboard for Viva Connections
    ```

=== "Markdown"

    ```md
    # spo web clientsidewebpart list --webUrl "https://reshmeeauckloo.sharepoint.com/sites/Company311"

    Date: 4/10/2023

    ## Dashboard for Viva Connections (9cc0f495-db64-4d74-b06b-a3de16231fe1)

    Property | Value
    ---------|-------
    Id | 9cc0f495-db64-4d74-b06b-a3de16231fe1
    Name | 9cc0f495-db64-4d74-b06b-a3de16231fe1
    Title | Dashboard for Viva Connections        
    ```
