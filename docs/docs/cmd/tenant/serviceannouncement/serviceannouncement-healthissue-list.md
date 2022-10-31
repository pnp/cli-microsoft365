# tenant serviceannouncement healthissue list

Gets all service health issues for the tenant.

## Usage

```sh
m365 tenant serviceannouncement healthissue list [options]
```

## Options

`-s, --service [service]`
: Retrieve service health issues for the particular service. If not provided, retrieves health issues for all services

--8<-- "docs/cmd/_global.md"

## Examples

Get service health issues of all services in Microsoft 365

```sh
m365 tenant serviceannouncement healthissue list
```

Get service health issues for Microsoft Forms

```sh
m365 tenant serviceannouncement healthissue list --service "Microsoft Forms"
```

## Response

=== "JSON"

    ```json
    [
      {
        "startDateTime": "2022-05-24T16:00:00Z",
        "endDateTime": "2022-05-24T22:20:00Z",
        "lastModifiedDateTime": "2022-05-24T22:27:18.63Z",
        "title": "Installation delays within the Power Platform admin center",
        "id": "CR384241",
        "impactDescription": "Users may have experienced delays when installing applications within the Power Platform admin center.",
        "classification": "advisory",
        "origin": "microsoft",
        "status": "serviceRestored",
        "service": "Dynamics 365 Apps",
        "feature": "Other",
        "featureGroup": "Other",
        "isResolved": true,
        "highImpact": null,
        "details": [],
        "posts": [
          {
            "createdDateTime": "2022-05-24T21:22:56.817Z",
            "postType": "regular",
            "description": {
              "contentType": "html",
              "content": "Title: Installation delays within the Power Platform admin center\\\n\nUser Impact: Users may experience delays when installing applications within the Power Platform admin center.\\\n\nWe are aware of an emerging issue where users are experiencing delays when installing applications through the Power Platform admin center. We are investigating the issue and will provide another update within the next 30 minutes.\\\n\nThis information is preliminary and may be subject to changes, corrections, and updates."
            }
          }
        ]
      }
    ]
    ```

=== "Text"

    ```text
    id        title
    --------  ----------------------------------------------------------
    CR384241  Installation delays within the Power Platform admin center
    ```

=== "CSV"

    ```csv
    id,title
    CR384241,Installation delays within the Power Platform admin center
    ```

## More information

- List serviceAnnouncement issues: [https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-issues](https://docs.microsoft.com/en-us/graph/api/serviceannouncement-list-issues)
