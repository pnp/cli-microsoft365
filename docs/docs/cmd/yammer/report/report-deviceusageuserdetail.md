# yammer report deviceusageuserdetail

Gets details about Yammer device usage by user

## Usage

```sh
m365 yammer report deviceusageuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about Yammer device usage by user for the last week

```sh
m365 yammer report deviceusageuserdetail --period D7
```

Gets details about Yammer device usage by user for July 1, 2019

```sh
m365 yammer report deviceusageuserdetail --date 2019-07-01
```

Gets details about Yammer device usage by user for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report deviceusageuserdetail --period D7 --output text > "deviceusageuserdetail.txt"
```

Gets details about Yammer device usage by user for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report deviceusageuserdetail --period D7 --output json > "deviceusageuserdetail.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-09",
        "User Principal Name": "0439A166C614C2E8C7B4075DC4752054",
        "Display Name": "2236A6E43D08F619FE695DF3B163A32F",
        "User State": "",
        "State Change Date": "",
        "Last Activity Date": "",
        "Used Web": "No",
        "Used Windows Phone": "No",
        "Used Android Phone": "No",
        "Used iPhone": "No",
        "Used iPad": "No",
        "Used Others": "No",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,User Principal Name,Display Name,User State,State Change Date,Last Activity Date,Used Web,Used Windows Phone,Used Android Phone,Used iPhone,Used iPad,Used Others,Report Period
    2022-11-09,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,,,,No,No,No,No,No,No,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,User Principal Name,Display Name,User State,State Change Date,Last Activity Date,Used Web,Used Windows Phone,Used Android Phone,Used iPhone,Used iPad,Used Others,Report Period
    2022-11-09,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,,,,No,No,No,No,No,No,7
    ```
