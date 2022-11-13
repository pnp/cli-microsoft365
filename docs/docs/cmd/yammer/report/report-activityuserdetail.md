# yammer report activityuserdetail

Gets details about Yammer activity by user

## Usage

```sh
m365 yammer report activityuserdetail [options]
```

## Options

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify the date or period, but not both.

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about Yammer activity by user for the last week

```sh
m365 yammer report activityuserdetail --period D7
```

Gets details about Yammer activity by user for May 1, 2019

```sh
m365 yammer report activityuserdetail --date 2019-05-01
```

Gets details about Yammer activity by user for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report activityuserdetail --period D7 --output text > "activityuserdetail.txt"
```

Gets details about Yammer activity by user for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report activityuserdetail --period D7 --output json > "activityuserdetail.json"
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
        "Posted Count": "0",
        "Read Count": "0",
        "Liked Count": "0",
        "Assigned Products": "MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,User Principal Name,Display Name,User State,State Change Date,Last Activity Date,Posted Count,Read Count,Liked Count,Assigned Products,Report Period
    2022-11-09,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,,,,0,0,0,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING),7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,User Principal Name,Display Name,User State,State Change Date,Last Activity Date,Posted Count,Read Count,Liked Count,Assigned Products,Report Period
    2022-11-09,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,,,,0,0,0,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING),7
    ```
