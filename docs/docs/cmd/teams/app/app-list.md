# teams app list

Lists apps from the Microsoft Teams app catalog or apps installed in the specified team

## Usage

```sh
m365 teams app list [options]
```

## Options

`-a, --all`
: Specify, to get apps from your organization only

`-i, --teamId [teamId]`
: The ID of the team for which to list installed apps. Specify either `teamId` or `teamName` but not both

`-t, --teamName [teamName]`
: The display name of the team for which to list installed apps. Specify either `teamId` or `teamName` but not both

--8<-- "docs/cmd/_global.md"

## Remarks

To list apps installed in the specified Microsoft Teams team, specify that team's ID using the `teamId` option. If the `teamId` option is not specified, the command will list apps available in the Teams app catalog.

## Examples

List all Microsoft Teams apps from your organization's app catalog only

```sh
m365 teams app list
```

List all apps from the Microsoft Teams app catalog and the Microsoft Teams store

```sh
m365 teams app list --all
```

List your organization's apps installed in the specified Microsoft Teams team with id _6f6fd3f7-9ba5-4488-bbe6-a789004d0d55_

```sh
m365 teams app list --teamId 6f6fd3f7-9ba5-4488-bbe6-a789004d0d55
```

List your organization's apps installed in the specified Microsoft Teams team with name _Team Name_

```sh
m365 teams app list --teamName "Team Name"
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
