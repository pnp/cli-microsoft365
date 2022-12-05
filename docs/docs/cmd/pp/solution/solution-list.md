# pp solution list

Lists solutions in a given environment.

## Usage

```sh
m365 pp solution list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment

`--asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List all solutions in a specific environment

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
```

List all solutions in a specific environment as Admin

```sh
m365 pp solution list --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --asAdmin
```

## Response

=== "JSON"

    ```json
    [
      {
        "solutionid": "00000001-0000-0000-0001-00000000009b",
        "uniquename": "Crc00f1",
        "version": "1.0.0.0",
        "installedon": "2021-10-01T21:54:14Z",
        "solutionpackageversion": null,
        "friendlyname": "Common Data Services Default Solution",
        "versionnumber": 860052,
        "publisherid": {
          "friendlyname": "CDS Default Publisher",
          "publisherid": "00000001-0000-0000-0000-00000000005a"
        }
      }
    ]
    ```

=== "Text"

    ```text
    uniquename  version    publisher
    ----------  ---------  ---------------------
    Crc00f1     1.0.0.0    CDS Default Publisher
    ```

=== "CSV"

    ```csv
    uniquename,version,publisher
    Crc00f1,1.0.0.0,CDS Default Publisher
    ```
