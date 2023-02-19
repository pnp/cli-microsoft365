# aad user license list

Lists the license details for a given user

## Usage

```sh
m365 aad user license list [options]
```

## Options

`--userId [userId]`
: The ID of the user. Specify either `userId` or `userName` but not both.

`--userName [userName]`
: User principal name of the user. Specify either `userId` or `userName` but not both.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    Don't specify any option to list license details of the current user.

## Examples

List license details of the current logged in user.

```sh
m365 aad user license list
```

List license details of a specific user by its UPN.

```sh
m365 aad user license list --userName john.doe@contoso.com
```

List license details of a specific user by its ID.

```sh
m365 aad user license list --userId 59f80e08-24b1-41f8-8586-16765fd830d3
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw",
        "skuId": "c42b9cae-ea4f-4ab7-9717-81576235ccac",
        "skuPartNumber": "DEVELOPERPACK_E5",
        "servicePlans": [
          {
            "servicePlanId": "b76fb638-6ba6-402a-b9f9-83d28acb3d86",
            "servicePlanName": "VIVA_LEARNING_SEEDED",
            "provisioningStatus": "PendingProvisioning",
            "appliesTo": "User"
          },
          {
            "servicePlanId": "7547a3fe-08ee-4ccb-b430-5077c5041653",
            "servicePlanName": "YAMMER_ENTERPRISE",
            "provisioningStatus": "Success",
            "appliesTo": "User"
          },
          {
            "servicePlanId": "eec0eb4f-6444-4f95-aba0-50c24d67f998",
            "servicePlanName": "AAD_PREMIUM_P2",
            "provisioningStatus": "Disabled",
            "appliesTo": "User"
          }
        ]
      }
    ]
    ```

=== "Text"

    ```text
    id           : x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw
    skuId        : c42b9cae-ea4f-4ab7-9717-81576235ccac
    skuPartNumber: DEVELOPERPACK_E5
    ```

=== "CSV"

    ```csv
    id,skuId,skuPartNumber
    x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw,c42b9cae-ea4f-4ab7-9717-81576235ccac,DEVELOPERPACK_E5
    ```

=== "Markdown"

    ```md
    # aad user license list --userId "0c9c625f-faa9-4c3b-8cd8-d874b869f78c"

    Date: 2/19/2023

    ## x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw

    Property | Value
    ---------|-------
    id | x4s03usaBkSMs5fbAhyttK6cK8RP6rdKlxeBV2I1zKw
    skuId | c42b9cae-ea4f-4ab7-9717-81576235ccac
    skuPartNumber | DEVELOPERPACK\_E5
    servicePlans | [{"servicePlanId":"b76fb638-6ba6-402a-b9f9-83d28acb3d86","servicePlanName":"VIVA\_LEARNING\_SEEDED","provisioningStatus":"PendingProvisioning","appliesTo":"User"},{"servicePlanId":"7547a3fe-08ee-4ccb-b430-5077c5041653","servicePlanName":"YAMMER\_ENTERPRISE","provisioningStatus":"Success","appliesTo":"User"},{"servicePlanId":"eec0eb4f-6444-4f95-aba0-50c24d67f998","servicePlanName":"AAD\_PREMIUM\_P2","provisioningStatus":"Disabled","appliesTo":"User"}]
    ```
