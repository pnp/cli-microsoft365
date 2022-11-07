# teams app publish

Publishes Teams app to the organization's app catalog

## Usage

```sh
m365 teams app publish [options]
```

## Options

`-p, --filePath <filePath>`
: Absolute or relative path to the Teams manifest zip file to add to the app catalog

--8<-- "docs/cmd/_global.md"

## Remarks

You can only publish a Teams app as a global administrator.

## Examples

Add the _teams-manifest.zip_ file to the organization's app catalog

```sh
m365 teams app publish --filePath ./teams-manifest.zip
```

## Response

=== "JSON"

    ```json
    {
        "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
        "externalId": "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
        "displayName": "Test App",
        "distributionMethod": "organization"
    }
    ```

=== "Text"

    ```text
    displayName       : Test App
    distributionMethod: organization
    externalId        : b5561ec9-8cab-4aa3-8aa2-d8d7172e4311
    id                : e3e29acb-8c79-412b-b746-e6c39ff4cd22
    ```

=== "CSV"

    ```csv
    id,externalId,displayName,distributionMethod
    e3e29acb-8c79-412b-b746-e6c39ff4cd22,b5561ec9-8cab-4aa3-8aa2-d8d7172e4311,Test App,organization
    ```
