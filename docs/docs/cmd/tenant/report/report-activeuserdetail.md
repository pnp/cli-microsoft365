# tenant report activeuserdetail

Gets details about Microsoft 365 active users

## Usage

```sh
m365 tenant report activeuserdetail [options]
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

Gets details about Microsoft 365 active users for the last week

```sh
m365 tenant report activeuserdetail --period D7
```

Gets details about Microsoft 365 active users for May 1, 2019

```sh
m365 tenant report activeuserdetail --date 2019-05-01
```

Gets details about Microsoft 365 active users for the last week and exports the report data in the specified path in text format

```sh
m365 tenant report activeuserdetail --period D7 --output text > "activeuserdetail.txt"
```

Gets details about Microsoft 365 active users for the last week and exports the report data in the specified path in json format

```sh
m365 tenant report activeuserdetail --period D7 --output json > "activeuserdetail.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-23",
        "User Principal Name": "0439A166C614C2E8C7B4075DC4752054",
        "Display Name": "2236A6E43D08F619FE695DF3B163A32F",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Has Exchange License": "True",
        "Has OneDrive License": "True",
        "Has SharePoint License": "True",
        "Has Skype For Business License": "True",
        "Has Yammer License": "True",
        "Has Teams License": "True",
        "Exchange Last Activity Date": "2020-03-27",
        "OneDrive Last Activity Date": "2020-03-27",
        "SharePoint Last Activity Date": "2020-04-30",
        "Skype For Business Last Activity Date": "2020-05-10",
        "Yammer Last Activity Date": "2020-05-10",
        "Teams Last Activity Date": "2020-05-10",
        "Exchange License Assign Date": "2020-02-26",
        "OneDrive License Assign Date": "2020-02-26",
        "SharePoint License Assign Date": "2020-02-26",
        "Skype For Business License Assign Date": "2020-02-26",
        "Yammer License Assign Date": "2020-02-26",
        "Teams License Assign Date": "2020-02-26",
        "Assigned Products": "MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Has Exchange License,Has OneDrive License,Has SharePoint License,Has Skype For Business License,Has Yammer License,Has Teams License,Exchange Last Activity Date,OneDrive Last Activity Date,SharePoint Last Activity Date,Skype For Business Last Activity Date,Yammer Last Activity Date,Teams Last Activity Date,Exchange License Assign Date,OneDrive License Assign Date,SharePoint License Assign Date,Skype For Business License Assign Date,Yammer License Assign Date,Teams License Assign Date,Assigned Products
    2022-10-23,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,False,,True,True,True,True,True,True,2020-09-12,2022-09-12,2021-10-30,2020-10-30,2019-04-21,2017-09-20,2021-01-10,2021-01-10,2021-01-10,2021-01-10,2021-01-10,2021-01-10,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)
    ```

=== "CSV"

    ```csv
    Report Refresh Date,User Principal Name,Display Name,Is Deleted,Deleted Date,Has Exchange License,Has OneDrive License,Has SharePoint License,Has Skype For Business License,Has Yammer License,Has Teams License,Exchange Last Activity Date,OneDrive Last Activity Date,SharePoint Last Activity Date,Skype For Business Last Activity Date,Yammer Last Activity Date,Teams Last Activity Date,Exchange License Assign Date,OneDrive License Assign Date,SharePoint License Assign Date,Skype For Business License Assign Date,Yammer License Assign Date,Teams License Assign Date,Assigned Products
    2022-10-23,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,False,,True,True,True,True,True,True,,2022-09-12,2020-09-12,2022-09-12,2021-10-30,2020-10-30,2019-04-21,2017-09-20,2021-01-10,2021-01-10,2021-01-10,2021-01-10,2021-01-10,2021-01-10,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)
    ```
