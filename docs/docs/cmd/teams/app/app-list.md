# teams app list

Lists apps from the Microsoft Teams app catalog

## Usage

```sh
m365 teams app list [options]
```

## Options

: `--distributionMethod`
: The distribution method for which apps must be listed. Allowed values `store, organization, sideloaded`.

--8<-- "docs/cmd/_global.md"

## Examples

List all apps from the Microsoft Teams app catalog

```sh
m365 teams app list
```

List all apps from the Microsoft Teams app catalog according to a given distribution method

```sh
m365 teams app list --distributionMethod 'store'
```

## Response

=== "JSON"

    ```json
    [
      {
        "id": "ffdb7239-3b58-46ba-b108-7f90a6d8799b",
        "externalId": null,
        "displayName": "Contoso App",
        "distributionMethod": "store"
      }
    ]
    ```

=== "Text"

    ```text
    id                                                        displayName                       distributionMethod
    --------------------------------------------------------  --------------------------------  ------------------
    ffdb7239-3b58-46ba-b108-7f90a6d8799b                      Contoso App                       store
    ```

=== "CSV"

    ```csv
    id,displayName,distributionMethod
    ffdb7239-3b58-46ba-b108-7f90a6d8799b,Contoso App,store
    ```
