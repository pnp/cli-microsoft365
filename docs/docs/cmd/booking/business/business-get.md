# booking business get

Retrieve the specified Microsoft Bookings business.

## Usage

```sh
m365 booking business get [options]
```

## Options

`-i, --id [id]`
: ID of the business. Specify either `id` or `name` but not both.

`-n, --name [name]`
: Name of the business. Specify either `id` or `name` but not both.

--8<-- "docs/cmd/_global.md"

## Examples

Retrieve the specified Microsoft Bookings business with the specified id.

```sh
m365 booking business get --id 'business@contoso.onmicrosoft.com'
```

Retrieve the specified Microsoft Bookings business with the specified name.

```sh
m365 booking business get --name 'business name'
```

## Response

=== "JSON"

    ```json
    {
      "id": "Accounting@contoso.onmicrosoft.com",
      "displayName": "Accounting",
      "businessType": "Financial services",
      "phone": "",
      "email": "runningdevadmin@contoso.onmicrosoft.com",
      "webSiteUrl": "",
      "defaultCurrencyIso": "USD",
      "isPublished": true,
      "publicUrl": "https://outlook.office365.com/owa/calendar/Accounting@contoso.onmicrosoft.com/bookings/",
      "languageTag": "",
      "address": {
        "street": "",
        "city": "",
        "state": "",
        "countryOrRegion": "",
        "postalCode": ""
      },
      "businessHours": [
        {
          "day": "monday",
          "timeSlots": [
            {
              "startTime": "08:00:00.0000000",
              "endTime": "17:00:00.0000000"
            }
          ]
        },
        {
          "day": "tuesday",
          "timeSlots": [
            {
              "startTime": "08:00:00.0000000",
              "endTime": "17:00:00.0000000"
            }
          ]
        },
        {
          "day": "wednesday",
          "timeSlots": [
            {
              "startTime": "08:00:00.0000000",
              "endTime": "17:00:00.0000000"
            }
          ]
        },
        {
          "day": "thursday",
          "timeSlots": [
            {
              "startTime": "08:00:00.0000000",
              "endTime": "17:00:00.0000000"
            }
          ]
        },
        {
          "day": "friday",
          "timeSlots": [
            {
              "startTime": "08:00:00.0000000",
              "endTime": "17:00:00.0000000"
            }
          ]
        },
        {
          "day": "saturday",
          "timeSlots": []
        },
        {
          "day": "sunday",
          "timeSlots": []
        }
      ],
      "schedulingPolicy": {
        "timeSlotInterval": "PT30M",
        "minimumLeadTime": "P1D",
        "maximumAdvance": "P365D",
        "sendConfirmationsToOwner": true,
        "allowStaffSelection": true
      }
    }
    ```

=== "Text"

    ```text
    businessType      : Financial services
    defaultCurrencyIso: USD
    displayName       : Accounting
    email             : runningdevadmin@contoso.onmicrosoft.com
    id                : Accounting@contoso.onmicrosoft.com
    phone             :
    ```

=== "CSV"

    ```csv
    id,displayName,businessType,phone,email,defaultCurrencyIso
    Accounting@contoso.onmicrosoft.com,Accounting,Financial services,,runningdevadmin@contoso.onmicrosoft.com,USD
    ```

=== "Markdown"

    ```md
    # booking business get --name "Accounting"

    Date: 5/29/2023

    ## Accounting (Accounting@contoso.onmicrosoft.com)

    Property | Value
    ---------|-------
    id | Accounting@contoso.onmicrosoft.com
    displayName | Accounting
    businessType | Financial services
    phone |
    email | runningdevadmin@contoso.onmicrosoft.com
    webSiteUrl |
    defaultCurrencyIso | USD
    isPublished | true
    publicUrl | https://outlook.office365.com/owa/calendar/Accounting@contoso.onmicrosoft.com/bookings/
    languageTag |
    ```
