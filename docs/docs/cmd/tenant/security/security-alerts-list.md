# tenant security alerts list

Gets the security alerts for a tenant

## Usage

```sh
m365 tenant security alerts list [options]
```

## Options

`--vendor [vendor]`
: The vendor to return alerts for. Possible values `Azure Advanced Threat Protection`, `Azure Security Center`, `Microsoft Cloud App Security`, `Azure Active Directory Identity Protection`, `Azure Sentinel`, `Microsoft Defender ATP`. If omitted, all alerts are returned

--8<-- "docs/cmd/_global.md"

## Examples

Get all security alerts for a tenant

```sh
m365 tenant security alerts list
```

Get security alerts for a vendor with name _Azure Sentinel_

```sh
m365 tenant security alerts list --vendor "Azure Sentinel"
```

## Response

=== "JSON"

    ```json
    [
      {
          "id": "2517536653831539999_658fa695-a5e6-4b60-ac7c-b2c1396df384",
          "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
          "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
          "riskScore": null,
          "tags": [],
          "activityGroupName": null,
          "assignedTo": null,
          "category": "Storage.Blob_GeoAnomaly",
          "closedDateTime": null,
          "comments": [],
          "confidence": null,
          "createdDateTime": "2022-03-30T13:19:15.8039138Z",
          "description": "Someone has accessed your Azure Storage account 'westeuropegivcekj' from an unusual location.",
          "detectionIds": [],
          "eventDateTime": "2022-03-30T10:16:56.846Z",
          "feedback": null,
          "incidentIds": [],
          "lastEventDateTime": null,
          "lastModifiedDateTime": "2022-03-30T13:19:48.5196488Z",
          "recommendedActions": [
            "• Limit access to your storage account, following the 'least privilege' principle: https://go.microsoft.com/fwlink/?linkid=2187303.• Consider using identity-based authentication: https://go.microsoft.com/fwlink/?linkid=2187202.• Consider rotating the storage account access keys and ensure that your access tokens are only shared with authorized users.• Ensure that storage access tokens are stored in a secured location such as Azure Key Vault. Avoid storing or sharing storage access tokens in source code, documentation, and email."
          ],
          "severity": "low",
          "sourceMaterials": [
            "https://portal.azure.com/#blade/Microsoft_Azure_Security_AzureDefenderForData/AlertBlade/alertId/2517536653831539999_658fa695-a5e6-4b60-ac7c-b2c1396df384/subscriptionId/bbdf91d0-d75b-430e-b738-2c525452144f/resourceGroup/managed-rg-purview-mip-poc/referencedFrom/alertDeepLink/location/westeurope"
          ],
          "status": "newAlert",
          "title": "Access from an unusual location to a storage blob container",
          "CustomProperties": "[\"{\\\"Alert Id\\\":\\\"658fa695-a5e6-4b60-ac7c-b2c1396df384\\\",\\\"Azure AD user\\\":\\\"N/A (Azure AD user authentication was not used)\\\",\\\"User agent\\\":\\\"Azure-Storage/9.3.0 (.NET Core)\\\",\\\"API type\\\":\\\"Blob\\\",\\\"Client location\\\":\\\"Dublin, Ireland\\\",\\\"Authentication type\\\":\\\"Shared access signature (SAS)\\\",\\\"Investigation steps\\\":\\\"{\\\\\\\"displayValue\\\\\\\":\\\\\\\"View related storage activity using Storage Analytics Logging. See how to configure Storage Analytics logging and more information\\\\\\\",\\\\\\\"kind\\\\\\\":\\\\\\\"Link\\\\\\\",\\\\\\\"value\\\\\\\":\\\\\\\"https:\\\\\\\\/\\\\\\\\/go.microsoft.com\\\\\\\\/fwlink\\\\\\\\/?linkid=2075734\\\\\\\"}\\\",\\\"Operations types\\\":\\\"GetBlob\\\",\\\"Service type\\\":\\\"Azure Blobs\\\",\\\"Container\\\":\\\"temporary\\\",\\\"Potential causes\\\":\\\"This alert indicates that this account has been accessed successfully from an IP address that is unfamiliar and unexpected compared to recent access pattern on this account.\\\\\\Potential causes:\\\\\\• An attacker has accessed your storage account.\\\\\\• A legitimate user has accessed your storage account from a new location.\\\",\\\"resourceType\\\":\\\"Storage\\\",\\\"ReportingSystem\\\":\\\"Azure\\\"}\",\"\\\"InitialAccess\\\"\"]",
          "vendorInformation": {
            "provider": "ASC",
            "providerVersion": null,
            "subProvider": "StorageThreatDetection",
            "vendor": "Microsoft"
          },
          "alertDetections": [],
          "cloudAppStates": [],
          "fileStates": [],
          "hostStates": [],
          "historyStates": [],
          "investigationSecurityStates": [],
          "malwareStates": [],
          "messageSecurityStates": [],
          "networkConnections": [
            {
              "applicationName": null,
              "destinationAddress": null,
              "destinationDomain": null,
              "destinationLocation": null,
              "destinationPort": null,
              "destinationUrl": null,
              "direction": null,
              "domainRegisteredDateTime": null,
              "localDnsName": null,
              "natDestinationAddress": null,
              "natDestinationPort": null,
              "natSourceAddress": null,
              "natSourcePort": null,
              "protocol": "tcp",
              "riskScore": null,
              "sourceAddress": "52.214.204.49",
              "sourceLocation": "Dublin, Dublin, IE",
              "sourcePort": null,
              "status": null,
              "urlParameters": null
            }
          ],
          "processes": [],
          "registryKeyStates": [],
          "securityResources": [
            {
              "resource": "/subscriptions/bbdf91d0-d75b-430e-b738-2c525452144f/resourceGroups/managed-rg-purview-mip-poc/providers/Microsoft.Storage/storageAccounts/scanwesteuropegivcebj",
              "resourceType": "attacked"
            }
          ],
          "triggers": [],
          "userStates": [],
          "uriClickSecurityStates": [],
          "vulnerabilityStates": []
        }
    ]
    ```

=== "Text"

    ```text
    id                                   title                      severity
    ------------------------------------ -------------------------- --------
    4ece2cf8-cbc0-5a42-92c3-e23f96006907 SharePoint Bulk Edit Items medium
    ```

=== "CSV"

    ```csv
    id,title,severity
    4ece2cf8-cbc0-5a42-92c3-e23f96006907,SharePoint Bulk Edit Items,medium
    ```

=== "Markdown"

    ```md
    # tenant security alerts list

    Date: 3/20/2022

    ## Unfamiliar sign-in properties (2517536653831539999_658fa695-a5e6-4b60-ac7c-b2c1396df384)

    Property | Value
    ---------|-------
    id | 2517536653831539999_658fa695-a5e6-4b60-ac7c-b2c1396df384
    azureTenantId | b8e1599d-b418-4be9-8f39-df03c3abe27a
    category | Storage.Blob_GeoAnomaly
    createdDateTime | 2022-03-30T13:19:15.8039138Z
    description | The following properties of this sign-in are unfamiliar for the given user: ASN, Browser, Device, IP, Location, EASId, TenantIPsubnet
    eventDateTime | 2022-03-30T10:16:56.846Z
    lastModifiedDateTime | 2022-03-30T13:19:48.5196488Z
    severity | low
    status | newAlert
    title | Access from an unusual location to a storage blob container
    ```
