# onedrive report usageaccountdetail

Gets details about OneDrive usage by account

## Usage

```sh
m365 onedrive report usageaccountdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets details about OneDrive usage by account for the last week

```sh
m365 onedrive report usageaccountdetail --period D7
```

Gets details about OneDrive usage by account for May 1, 2019

```sh
m365 onedrive report usageaccountdetail --date 2019-05-01
```

Gets details about OneDrive usage by account for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report usageaccountdetail --period D7 --output text > "onedriveusageaccountdetail.txt"
```

Gets details about OneDrive usage by account for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report usageaccountdetail --period D7 --output json > "onedriveusageaccountdetail.json"
```

## Response

=== "JSON"

```json
[
  {
    "Report Refresh Date": "2022-10-25",
    "Site URL": "C21296B9EFDC43D565AF865B09EAD4CB",
    "Owner Display Name": "C5A83451EDF152DF3556E14B13197438",
    "Is Deleted": "False",
    "Last Activity Date": "2022-10-13",
    "File Count": "87",
    "Active File Count": "0",
    "Storage Used (Byte)": "355367152",
    "Storage Allocated (Byte)": "1099511627776",
    "Owner Principal Name": "CD2A821A5A931201567E1EA5638A52FE",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Site URL": "C06B5B03DD0DB7D8DD0534010CFA7409",
    "Owner Display Name": "C523C55CE6690474842BFF7A0FC771CA",
    "Is Deleted": "False",
    "Last Activity Date": "2022-08-10",
    "File Count": "68",
    "Active File Count": "0",
    "Storage Used (Byte)": "25448429",
    "Storage Allocated (Byte)": "1099511627776",
    "Owner Principal Name": "C93FE6370AF6AB1FF720026C49DB4B11",
    "Report Period": "7"
  }
]
```

=== "Text"

    ``` text

Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period
2022-10-25,C21296B9EFDC43D565AF865B09EAD4CB,C5A83451EDF152DF3556E14B13197438,False,2022-10-13,87,0,355367152,1099511627776,CD2A821A5A931201567E1EA5638A52FE,7
2022-10-25,C06B5B03DD0DB7D8DD0534010CFA7409,C523C55CE6690474842BFF7A0FC771CA,False,2022-08-10,68,0,25448429,1099511627776,C93FE6370AF6AB1FF720026C49DB4B11,7

````

=== "CSV"

    ``` text
Report Refresh Date,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Storage Used (Byte),Storage Allocated (Byte),Owner Principal Name,Report Period
2022-10-25,C21296B9EFDC43D565AF865B09EAD4CB,C5A83451EDF152DF3556E14B13197438,False,2022-10-13,87,0,355367152,1099511627776,CD2A821A5A931201567E1EA5638A52FE,7
2022-10-25,C06B5B03DD0DB7D8DD0534010CFA7409,C523C55CE6690474842BFF7A0FC771CA,False,2022-08-10,68,0,25448429,1099511627776,C93FE6370AF6AB1FF720026C49DB4B11,7
````
