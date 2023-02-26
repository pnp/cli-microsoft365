# aad license list

Lists commercial subscriptions that an organization has acquired

## Usage

```sh
m365 aad license list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

List all licenses within the tenant

```sh
m365 aad license list
```

## Response

=== "JSON"

    ```json
    [
      {
        "capabilityStatus": "Enabled",
        "consumedUnits": 1,
        "id": "48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df",
        "skuId": "c7df2760-2c81-4ef7-b578-5b5392b571df",
        "skuPartNumber": "ENTERPRISEPREMIUM",
        "appliesTo": "User",
        "prepaidUnits": {
          "enabled": 10000,
          "suspended": 0,
          "warning": 0
        },
        "servicePlans": [
            {
                "servicePlanId": "8c098270-9dd4-4350-9b30-ba4703f3b36b",
                "servicePlanName": "ADALLOM_S_O365",
                "provisioningStatus": "Success",
                "appliesTo": "User"
            }
        ]
      }
    ]
    ```

=== "Text"

    ```text
    id                                                                         skuId                                 skuPartNumber
    -------------------------------------------------------------------------  ------------------------------------  ----------------------
    48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df  c7df2760-2c81-4ef7-b578-5b5392b571df  ENTERPRISEPREMIUM
    ```

=== "CSV"

    ```csv
    id,skuId,skuPartNumber
    48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df,c7df2760-2c81-4ef7-b578-5b5392b571df,ENTERPRISEPREMIUM
    ```
    
=== "Markdown"

    ```md
    # aad license list

    Date: 14/2/2023

    ## 48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df

    Property | Value
    ---------|-------
    capabilityStatus | Enabled
    consumedUnits | 1
    id | 48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df
    skuId | c7df2760-2c81-4ef7-b578-5b5392b571df
    skuPartNumber | ENTERPRISEPREMIUM
    appliesTo | User
    prepaidUnits | {"enabled":10000,"suspended":0,"warning":0}
    servicePlans | [{"servicePlanId":"8c098270-9dd4-4350-9b30-ba4703f3b36b","servicePlanName": "ADALLOM_S_O365","provisioningStatus": "Success","appliesTo": "User"}]
    ```
