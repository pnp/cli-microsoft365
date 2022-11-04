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
    Accounting@8b7jz7.onmicrosoft.com  Accounting
    BestShop@8b7jz7.onmicrosoft.com    Best Shop
    ```

=== "CSV"

    ```csv
    id,displayName
    Accounting@8b7jz7.onmicrosoft.com,Accounting
    BestShop@8b7jz7.onmicrosoft.com,Best Shop
    ```
