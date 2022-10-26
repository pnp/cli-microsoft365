# teams report deviceusageuserdetail

Gets detail about Microsoft Teams device usage by user.

## Usage

```sh
m365 teams report deviceusageuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`.

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams device usage by user report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets information about Microsoft Teams device usage by user for the last week

```sh
m365 teams report deviceusageuserdetail --period D7
```

Gets information about Microsoft Teams device usage by user for July 1, 2019

```sh
m365 teams report deviceusageuserdetail --date 2019-07-01
```

Gets information about Microsoft Teams device usage by user for the last week and exports the report data in the specified path in text format

```sh
m365 teams report deviceusageuserdetail --period D7 --output text > "deviceusageuserdetail.txt"
```

Gets information about Microsoft Teams device usage by user for the last week and exports the report data in the specified path in json format

```sh
m365 teams report deviceusageuserdetail --period D7 --output json > "deviceusageuserdetail.json"
```

## Response

=== "JSON"

    ``` json
    [
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "77E5979DD60BA6EAA53E814DBEEEFA5F",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "3A50C39C57095E612D2E859BDC88F3DE",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "F29840C36A2FABA7A31122484414D11D",
        "Last Activity Date": "2022-07-18",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "4416DC264370E7E989FEF6BE4E63FBC3",
        "Last Activity Date": "2022-06-17",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "0748A206D73F08EF088BD021B8300D97",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "66102E4089CB316D823BCAF99DFC5371",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "B216325CF128EC5F7B525D1B55F68414",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "9767FDB2F335B6DC60B45AFC96A2F2DE",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "0486161B72ED086483ED9515BAEFBB9D",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "04A2BBB23350895E2B5A0D3B8F00A2DF",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "ED6BC44459F5CE821E73903442D134BE",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "457447629A0ECC9CB43E37A35E65AD8E",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "7517646B39D4003BDA6FC54757E4F5EC",
        "Last Activity Date": "2022-10-24",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "Yes",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "362626EFDBED0F0A3971FF1318673AAB",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "45C95E840B021CD8B07EFC0A8F223282",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "7F7855C9A156C2239B955F796D723D29",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "AD598B02C0E72C67E5030DDE0B9444C0",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "4494A93CA6E2415AB0F8B8BBDBC96350",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "626081C403EF53E9664D83AD14D59F2D",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "5AAD23A07D61D4B09F102E87AE00F7E2",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "554630B7593DDE8E04F27933A965D5B2",
        "Last Activity Date": "2022-10-11",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "54A5C6D2A26D2F003F2D3C96800EEAE1",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "A8E64B7615C6F55588F8B4EC5AC4EB68",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-24",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "0439A166C614C2E8C7B4075DC4752054",
        "Last Activity Date": "",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used iOS": "No",
        "Used Mac": "No",
        "Used Android Phone": "No",
        "Used Windows": "No",
        "Used  Chrome OS": "No",
        "Used Linux": "No",
        "Is Licensed": "Yes",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ``` text
    Report Refresh Date,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Used Web,Used Windows Phone,Used iOS,Used Mac,Used Android Phone,Used Windows,Used  Chrome OS,Used Linux,Is Licensed,Report Period
    2022-10-24,00000000-0000-0000-0000-000000000000,77E5979DD60BA6EAA53E814DBEEEFA5F,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,3A50C39C57095E612D2E859BDC88F3DE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,F29840C36A2FABA7A31122484414D11D,2022-07-18,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,4416DC264370E7E989FEF6BE4E63FBC3,2022-06-17,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0748A206D73F08EF088BD021B8300D97,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,66102E4089CB316D823BCAF99DFC5371,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,B216325CF128EC5F7B525D1B55F68414,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,9767FDB2F335B6DC60B45AFC96A2F2DE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0486161B72ED086483ED9515BAEFBB9D,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,04A2BBB23350895E2B5A0D3B8F00A2DF,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,ED6BC44459F5CE821E73903442D134BE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,457447629A0ECC9CB43E37A35E65AD8E,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,7517646B39D4003BDA6FC54757E4F5EC,2022-10-24,False,,No,No,No,No,No,Yes,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,362626EFDBED0F0A3971FF1318673AAB,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,45C95E840B021CD8B07EFC0A8F223282,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,7F7855C9A156C2239B955F796D723D29,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,AD598B02C0E72C67E5030DDE0B9444C0,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,4494A93CA6E2415AB0F8B8BBDBC96350,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,626081C403EF53E9664D83AD14D59F2D,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,5AAD23A07D61D4B09F102E87AE00F7E2,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,554630B7593DDE8E04F27933A965D5B2,2022-10-11,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,54A5C6D2A26D2F003F2D3C96800EEAE1,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,A8E64B7615C6F55588F8B4EC5AC4EB68,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0439A166C614C2E8C7B4075DC4752054,,False,,No,No,No,No,No,No,No,No,Yes,7
    ```

=== "CSV"

    ``` text
    Report Refresh Date,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Used Web,Used Windows Phone,Used iOS,Used Mac,Used Android Phone,Used Windows,Used  Chrome OS,Used Linux,Is Licensed,Report Period
    2022-10-24,00000000-0000-0000-0000-000000000000,77E5979DD60BA6EAA53E814DBEEEFA5F,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,3A50C39C57095E612D2E859BDC88F3DE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,F29840C36A2FABA7A31122484414D11D,2022-07-18,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,4416DC264370E7E989FEF6BE4E63FBC3,2022-06-17,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0748A206D73F08EF088BD021B8300D97,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,66102E4089CB316D823BCAF99DFC5371,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,B216325CF128EC5F7B525D1B55F68414,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,9767FDB2F335B6DC60B45AFC96A2F2DE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0486161B72ED086483ED9515BAEFBB9D,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,04A2BBB23350895E2B5A0D3B8F00A2DF,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,ED6BC44459F5CE821E73903442D134BE,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,457447629A0ECC9CB43E37A35E65AD8E,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,7517646B39D4003BDA6FC54757E4F5EC,2022-10-24,False,,No,No,No,No,No,Yes,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,362626EFDBED0F0A3971FF1318673AAB,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,45C95E840B021CD8B07EFC0A8F223282,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,7F7855C9A156C2239B955F796D723D29,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,AD598B02C0E72C67E5030DDE0B9444C0,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,4494A93CA6E2415AB0F8B8BBDBC96350,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,626081C403EF53E9664D83AD14D59F2D,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,5AAD23A07D61D4B09F102E87AE00F7E2,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,554630B7593DDE8E04F27933A965D5B2,2022-10-11,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,54A5C6D2A26D2F003F2D3C96800EEAE1,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,A8E64B7615C6F55588F8B4EC5AC4EB68,,False,,No,No,No,No,No,No,No,No,Yes,7
    2022-10-24,00000000-0000-0000-0000-000000000000,0439A166C614C2E8C7B4075DC4752054,,False,,No,No,No,No,No,No,No,No,Yes,7
    ```
