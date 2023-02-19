# teams report pstncalls

Get details about PSTN calls made within a given time period

## Usage

```sh
m365 teams report pstncalls [options]
```

## Options

`--fromDateTime <fromDateTime>`
: The start of time range to query. UTC, inclusive

`--toDateTime [toDateTime]`
: The end time range to query. UTC, inclusive. Defaults to today if omitted

--8<-- "docs/cmd/_global.md"

## Remarks

This command only works with app-only permissions. You will need to create your own Azure AD app with `CallRecords.Read.All` permission assigned. Instructions on how to create your own Azure AD app can be found at [Using your own Azure AD identity](../../../user-guide/using-own-identity.md)

The difference between `fromDateTime` and `toDateTime` cannot exceed a period of 90 days

## Examples

Get details about PSTN calls made between 2020-10-31 and today

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in text format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output text > "pstncalls.txt"
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in json format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output json > "pstncalls.json"
```

## More information

- List PSTN calls: [https://docs.microsoft.com/en-us/graph/api/callrecords-callrecord-getpstncalls?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/callrecords-callrecord-getpstncalls?view=graph-rest-1.0)

## Response

=== "JSON"

    ``` json
    {
      "@odata.count": 1,
      "value": [
        {
          "id": "9c4984c7-6c3c-427d-a30c-bd0b2eacee90",
          "callId": "1835317186_112562680@61.221.3.176",
          "userId": "db03c14b-06eb-4189-939b-7cbf3a20ba27",
          "userPrincipalName": "richard.malk@contoso.com",
          "userDisplayName": "Richard Malk",
          "startDateTime": "2019-11-01T00:00:08.2589935Z",
          "endDateTime": "2019-11-01T00:03:47.2589935Z",
          "duration": 219,
          "charge": 0.00,
          "callType": "user_in",
          "currency": "USD",
          "calleeNumber": "+1234567890",
          "usageCountryCode": "US",
          "tenantCountryCode": "US",
          "connectionCharge": 0.00,
          "callerNumber": "+0123456789",
          "destinationContext": null,
          "destinationName": "United States",
          "conferenceId": null,
          "licenseCapability": "MCOPSTNU",
          "inventoryType": "Subscriber"
        }
      ]
  }
    ```

=== "Text"

    ``` text
    id,calleeNumber,callerNumber,startDateTime
    9c4984c7-6c3c-427d-a30c-bd0b2eacee90,+1234567890,+0123456789,2019-11-01T00:00:08.2589935Z
    ```

=== "CSV"

    ``` text
    id,calleeNumber,callerNumber,startDateTime
    9c4984c7-6c3c-427d-a30c-bd0b2eacee90,+1234567890,+0123456789,2019-11-01T00:00:08.2589935Z
    ```
