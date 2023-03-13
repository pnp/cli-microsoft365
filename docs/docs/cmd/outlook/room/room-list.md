# outlook room list

Get a collection of all available rooms

## Usage

```sh
m365 outlook room list [options]
```

## Options

`--roomlistEmail [roomlistEmail]`
: Use to filter returned rooms by their roomlist email (eg. bldg2@contoso.com)

--8<-- "docs/cmd/_global.md"

## Examples

Get all the rooms

```sh
m365 outlook room list
```

Get all the rooms of specified roomlist e-mail address

```sh
m365 outlook room list --roomlistEmail "bldg2@contoso.com"
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "98c40767-158a-44f0-9dda-c95b86f079ca",
        "emailAddress": "Largeroom@contoso.com",
        "displayName": "Large room",
        "geoCoordinates": null,
        "phone": "",
        "nickname": "Large room",
        "building": null,
        "floorNumber": null,
        "floorLabel": null,
        "label": null,
        "capacity": 25,
        "bookingType": "standard",
        "audioDeviceName": null,
        "videoDeviceName": null,
        "displayDeviceName": null,
        "isWheelChairAccessible": false,
        "tags": [],
        "address": {
          "street": "Microsoft Way 1",
          "city": "Redmond",
          "state": "Washington",
          "countryOrRegion": "US",
          "postalCode": "98053"
        }
      }
    ]
    ```

=== "Text"

    ```txt
    id                                    displayName  phone  emailAddress
    ------------------------------------  -----------  -----  ---------------------
    98c40767-158a-44f0-9dda-c95b86f079ca  Large room          Largeroom@contoso.com
    ```

=== "CSV"

    ```csv
    id,displayName,phone,emailAddress
    98c40767-158a-44f0-9dda-c95b86f079ca,Large room,,Largeroom@contoso.com
    ```

=== "Markdown"

    ```md
    # outlook room list

    Date: 27/1/2023

    ## Large room (98c40767-158a-44f0-9dda-c95b86f079ca)

    Property | Value
    ---------|-------
    id | 98c40767-158a-44f0-9dda-c95b86f079ca
    emailAddress | Largeroom@VanRoeyBeSPDev.onmicrosoft.com
    displayName | Large room
    geoCoordinates | null
    phone |
    nickname | Large room
    building | null
    floorNumber | null
    floorLabel | null
    label | null
    capacity | 25
    bookingType | standard
    audioDeviceName | null
    videoDeviceName | null
    displayDeviceName | null
    isWheelChairAccessible | false
    tags | []
    address | {"street":"Microsoft Way 1","city":"Redmond","state":"Washington","countryOrRegion":"US","postalCode":"98053"}
    ```
