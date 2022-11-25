# onedrive report usagestorage

Gets the trend on the amount of storage you are using in OneDrive for Business

## Usage

```sh
m365 onedrive report usagestorage [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week

```sh
m365 onedrive report usagestorage --period D7
```

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report usagestorage --period D7 --output text > "usagestorage.txt"
```

Gets the trend on the amount of storage you are using in OneDrive for Business for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report usagestorage --period D7 --output json > "usagestorage.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2079158662703",
        "Report Date": "2022-10-25",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2079158662703",
        "Report Date": "2022-10-25",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2079158662703",
        "Report Date": "2022-10-24",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2079158662703",
        "Report Date": "2022-10-24",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2079174134177",
        "Report Date": "2022-10-23",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2079174134177",
        "Report Date": "2022-10-23",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2078145067718",
        "Report Date": "2022-10-22",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2078145067718",
        "Report Date": "2022-10-22",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2070117199614",
        "Report Date": "2022-10-21",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2070117199614",
        "Report Date": "2022-10-21",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2070117199614",
        "Report Date": "2022-10-20",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2070117199614",
        "Report Date": "2022-10-20",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "OneDrive",
        "Storage Used (Byte)": "2069382310800",
        "Report Date": "2022-10-19",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Storage Used (Byte)": "2069382310800",
        "Report Date": "2022-10-19",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```csv
    Report Refresh Date,Site Type,Storage Used (Byte),Report Date,Report Period
    2022-10-25,OneDrive,2079158662703,2022-10-25,7
    2022-10-25,All,2079158662703,2022-10-25,7
    2022-10-25,OneDrive,2079158662703,2022-10-24,7
    2022-10-25,All,2079158662703,2022-10-24,7
    2022-10-25,OneDrive,2079174134177,2022-10-23,7
    2022-10-25,All,2079174134177,2022-10-23,7
    2022-10-25,OneDrive,2078145067718,2022-10-22,7
    2022-10-25,All,2078145067718,2022-10-22,7
    2022-10-25,OneDrive,2070117199614,2022-10-21,7
    2022-10-25,All,2070117199614,2022-10-21,7
    2022-10-25,OneDrive,2070117199614,2022-10-20,7
    2022-10-25,All,2070117199614,2022-10-20,7
    2022-10-25,OneDrive,2069382310800,2022-10-19,7
    2022-10-25,All,2069382310800,2022-10-19,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Storage Used (Byte),Report Date,Report Period
    2022-10-25,OneDrive,2079158662703,2022-10-25,7
    2022-10-25,All,2079158662703,2022-10-25,7
    2022-10-25,OneDrive,2079158662703,2022-10-24,7
    2022-10-25,All,2079158662703,2022-10-24,7
    2022-10-25,OneDrive,2079174134177,2022-10-23,7
    2022-10-25,All,2079174134177,2022-10-23,7
    2022-10-25,OneDrive,2078145067718,2022-10-22,7
    2022-10-25,All,2078145067718,2022-10-22,7
    2022-10-25,OneDrive,2070117199614,2022-10-21,7
    2022-10-25,All,2070117199614,2022-10-21,7
    2022-10-25,OneDrive,2070117199614,2022-10-20,7
    2022-10-25,All,2070117199614,2022-10-20,7
    2022-10-25,OneDrive,2069382310800,2022-10-19,7
    2022-10-25,All,2069382310800,2022-10-19,7
    ```
