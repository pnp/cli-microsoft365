# booking business list

Lists all Microsoft Bookings businesses that are created for the tenant.

## Usage

```sh
m365 booking business list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Returns a list of all Microsoft Bookings businesses that are created for the tenant.

```sh
m365 booking business list
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "Accounting@8b7jz1.onmicrosoft.com",
        "displayName": "Accounting"
      },
      {
        "id": "BestShop@8b7jz1.onmicrosoft.com",
        "displayName": "Best Shop"
      }
    ]
    ```

=== "Text"

    ```text
    id                                 displayName
    ---------------------------------  -----------
    Accounting@contoso.onmicrosoft.com  Accounting
    BestShop@contoso.onmicrosoft.com    Best Shop
    ```

=== "CSV"

    ```csv
    id,displayName
    Accounting@contoso.onmicrosoft.com,Accounting
    BestShop@contoso.onmicrosoft.com,Best Shop
    ```

=== "Markdown"

    ```md
    # booking business list

    Date: 5/29/2023

    ## Accounting (Accounting@contoso.onmicrosoft.com)

    Property | Value
    ---------|-------
    id | Accounting@contoso.onmicrosoft.com
    displayName | Accounting

    ## Best Shop (BestShop@contoso.onmicrosoft.com)

    Property | Value
    ---------|-------
    id | BestShop@contoso.onmicrosoft.com
    displayName | Best Shop
    ```
