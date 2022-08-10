import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./security-alerts-list');

describe(commands.SECURITY_ALERTS_LIST, () => {

  const alertResponseCSV = `id,title,severity
  4ece2cf8-cbc0-5a42-92c3-e23f96006907,SharePoint Bulk Edit Items,medium
  33aed7062fce896e48e2f63fe3971153b0bb959a3ac25fd3b282c469b2cb54a7,Anonymous IP address,medium
  6254ad90467a7d7b3d69f934,Multiple failed login attempts,low`;

  const alertASC = {
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
  };

  const alertMCAS = {
    "id": "6254ad90467a7d7b3d69f934",
    "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
    "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
    "riskScore": null,
    "tags": [],
    "activityGroupName": null,
    "assignedTo": null,
    "category": "MCAS_ALERT_ANUBIS_DETECTION_REPEATED_ACTIVITY_FAILED_LOGIN",
    "closedDateTime": null,
    "comments": [],
    "confidence": null,
    "createdDateTime": "2022-04-11T22:37:04Z",
    "description": "The user \"jondoe (jondoe@contoso.onmicrosoft.com)\" performed more than 103 failed logins attempts in a single session.",
    "detectionIds": [],
    "eventDateTime": "2022-04-11T07:22:39.69Z",
    "feedback": null,
    "incidentIds": [],
    "lastEventDateTime": null,
    "lastModifiedDateTime": "2022-04-11T22:37:05.5976027Z",
    "recommendedActions": [],
    "severity": "low",
    "sourceMaterials": [
      "https://contoso.portal.cloudappsecurity.com/#/alerts/6254ad90467a7d7b3d69f934",
      "https://contoso.portal.cloudappsecurity.com/#/policy/?id=eq(5a4cd40810bc95f3d6cbaa83,)",
      "https://contoso.portal.cloudappsecurity.com/#/alerts/6254ad90467a7d7b3d69f934"
    ],
    "status": "newAlert",
    "title": "Multiple failed login attempts",
    "vendorInformation": {
      "provider": "MCAS",
      "providerVersion": null,
      "subProvider": null,
      "vendor": "Microsoft"
    },
    "alertDetections": [],
    "cloudAppStates": [
      {
        "destinationServiceIp": null,
        "destinationServiceName": "Microsoft SharePoint Online",
        "riskScore": null
      },
      {
        "destinationServiceIp": null,
        "destinationServiceName": "Office 365",
        "riskScore": null
      }
    ],
    "fileStates": [],
    "hostStates": [],
    "historyStates": [],
    "investigationSecurityStates": [],
    "malwareStates": [],
    "messageSecurityStates": [],
    "networkConnections": [],
    "processes": [],
    "registryKeyStates": [],
    "securityResources": [],
    "triggers": [],
    "userStates": [
      {
        "aadUserId": "b9e36c12-4683-4da9-bf7d-d14f73e7bd2c",
        "accountName": "jondoe",
        "domainName": "contoso.onmicrosoft.com",
        "emailRole": "unknown",
        "isVpn": null,
        "logonDateTime": null,
        "logonId": null,
        "logonIp": null,
        "logonLocation": null,
        "logonType": null,
        "onPremisesSecurityIdentifier": null,
        "riskScore": null,
        "userAccountType": null,
        "userPrincipalName": "jondoe@contoso.onmicrosoft.com"
      }
    ],
    "uriClickSecurityStates": [],
    "vulnerabilityStates": []
  };
  
  const alertIPC = {
    "id": "33aed7062fce896e48e2f63fe3971153b0bb959a3ac25fd3b282c469b2cb54a7",
    "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
    "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
    "riskScore": null,
    "tags": [],
    "activityGroupName": null,
    "assignedTo": null,
    "category": "AnonymousLogin",
    "closedDateTime": null,
    "comments": [],
    "confidence": null,
    "createdDateTime": "2022-04-12T01:49:42.3797106Z",
    "description": "Sign-in from an anonymous IP address (e.g. Tor browser, anonymizer VPNs)",
    "detectionIds": [],
    "eventDateTime": "2022-04-12T01:49:42.3797106Z",
    "feedback": null,
    "incidentIds": [],
    "lastEventDateTime": null,
    "lastModifiedDateTime": "2022-04-12T01:51:14.5996188Z",
    "recommendedActions": [],
    "severity": "medium",
    "sourceMaterials": [],
    "status": "newAlert",
    "title": "Anonymous IP address",
    "vendorInformation": {
      "provider": "IPC",
      "providerVersion": null,
      "subProvider": null,
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
    "networkConnections": [],
    "processes": [],
    "registryKeyStates": [],
    "securityResources": [],
    "triggers": [],
    "userStates": [
      {
        "aadUserId": "0b0274d6-0398-4026-a742-03be2dd5f440",
        "accountName": "jon.doe",
        "domainName": "contoso.onmicrosoft.com",
        "emailRole": "unknown",
        "isVpn": null,
        "logonDateTime": "2022-04-12T01:49:42.3797106Z",
        "logonId": null,
        "logonIp": "183.220.101.28",
        "logonLocation": "Schoenwalde-Glien, Brandenburg, DE",
        "logonType": null,
        "onPremisesSecurityIdentifier": null,
        "riskScore": null,
        "userAccountType": null,
        "userPrincipalName": "jon.doe@contoso.onmicrosoft.com"
      }
    ],
    "uriClickSecurityStates": [],
    "vulnerabilityStates": []
  };

  const alertResponse = [
    {
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#security/alerts",
      "value": [
        {
          "id": "4ece2cf8-cbc0-5a42-92c3-e23f96006907",
          "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
          "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
          "riskScore": null,
          "tags": [],
          "activityGroupName": null,
          "assignedTo": null,
          "category": "5e462672-358b-48cb-9ca9-c910c99cb34d_0eb2a7bf-4d08-469a-8344-267c7b749779",
          "closedDateTime": null,
          "comments": [],
          "confidence": null,
          "createdDateTime": "2022-04-11T17:03:34.0185371Z",
          "description": "Detect bulk edits in SharePoint",
          "detectionIds": [],
          "eventDateTime": "2022-04-11T12:52:50Z",
          "feedback": null,
          "incidentIds": [],
          "lastEventDateTime": null,
          "lastModifiedDateTime": "2022-04-11T17:03:34.4787699Z",
          "recommendedActions": [],
          "severity": "medium",
          "sourceMaterials": [],
          "status": "newAlert",
          "title": "SharePoint Bulk Edit Items",
          "vendorInformation": {
            "provider": "Azure Sentinel",
            "providerVersion": null,
            "subProvider": null,
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
          "networkConnections": [],
          "processes": [],
          "registryKeyStates": [],
          "securityResources": [],
          "triggers": [],
          "userStates": [],
          "uriClickSecurityStates": [],
          "vulnerabilityStates": []
        },
        {
          "id": "33aed7062fce896e48e2f63fe3971153b0bb959a3ac25fd3b282c469b2cb54a7",
          "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
          "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
          "riskScore": null,
          "tags": [],
          "activityGroupName": null,
          "assignedTo": null,
          "category": "AnonymousLogin",
          "closedDateTime": null,
          "comments": [],
          "confidence": null,
          "createdDateTime": "2022-04-12T01:49:42.3797106Z",
          "description": "Sign-in from an anonymous IP address (e.g. Tor browser, anonymizer VPNs)",
          "detectionIds": [],
          "eventDateTime": "2022-04-12T01:49:42.3797106Z",
          "feedback": null,
          "incidentIds": [],
          "lastEventDateTime": null,
          "lastModifiedDateTime": "2022-04-12T01:51:14.5996188Z",
          "recommendedActions": [],
          "severity": "medium",
          "sourceMaterials": [],
          "status": "newAlert",
          "title": "Anonymous IP address",
          "vendorInformation": {
            "provider": "IPC",
            "providerVersion": null,
            "subProvider": null,
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
          "networkConnections": [],
          "processes": [],
          "registryKeyStates": [],
          "securityResources": [],
          "triggers": [],
          "userStates": [
            {
              "aadUserId": "0b0274d6-0398-4026-a742-03be2dd5f440",
              "accountName": "jon.doe",
              "domainName": "contoso.onmicrosoft.com",
              "emailRole": "unknown",
              "isVpn": null,
              "logonDateTime": "2022-04-12T01:49:42.3797106Z",
              "logonId": null,
              "logonIp": "183.220.101.28",
              "logonLocation": "Schoenwalde-Glien, Brandenburg, DE",
              "logonType": null,
              "onPremisesSecurityIdentifier": null,
              "riskScore": null,
              "userAccountType": null,
              "userPrincipalName": "jon.doe@contoso.onmicrosoft.com"
            }
          ],
          "uriClickSecurityStates": [],
          "vulnerabilityStates": []
        },
        {
          "id": "6254ad90467a7d7b3d69f934",
          "azureTenantId": "b8e1599d-b418-4be9-8f39-df03c3abe27a",
          "azureSubscriptionId": "ee390228-e284-4e54-a464-d693a1d55540",
          "riskScore": null,
          "tags": [],
          "activityGroupName": null,
          "assignedTo": null,
          "category": "MCAS_ALERT_ANUBIS_DETECTION_REPEATED_ACTIVITY_FAILED_LOGIN",
          "closedDateTime": null,
          "comments": [],
          "confidence": null,
          "createdDateTime": "2022-04-11T22:37:04Z",
          "description": "The user \"jondoe (jondoe@contoso.onmicrosoft.com)\" performed more than 103 failed logins attempts in a single session.",
          "detectionIds": [],
          "eventDateTime": "2022-04-11T07:22:39.69Z",
          "feedback": null,
          "incidentIds": [],
          "lastEventDateTime": null,
          "lastModifiedDateTime": "2022-04-11T22:37:05.5976027Z",
          "recommendedActions": [],
          "severity": "low",
          "sourceMaterials": [
            "https://contoso.portal.cloudappsecurity.com/#/alerts/6254ad90467a7d7b3d69f934",
            "https://contoso.portal.cloudappsecurity.com/#/policy/?id=eq(5a4cd40810bc95f3d6cbaa83,)",
            "https://contoso.portal.cloudappsecurity.com/#/alerts/6254ad90467a7d7b3d69f934"
          ],
          "status": "newAlert",
          "title": "Multiple failed login attempts",
          "vendorInformation": {
            "provider": "MCAS",
            "providerVersion": null,
            "subProvider": null,
            "vendor": "Microsoft"
          },
          "alertDetections": [],
          "cloudAppStates": [
            {
              "destinationServiceIp": null,
              "destinationServiceName": "Microsoft SharePoint Online",
              "riskScore": null
            },
            {
              "destinationServiceIp": null,
              "destinationServiceName": "Office 365",
              "riskScore": null
            }
          ],
          "fileStates": [],
          "hostStates": [],
          "historyStates": [],
          "investigationSecurityStates": [],
          "malwareStates": [],
          "messageSecurityStates": [],
          "networkConnections": [],
          "processes": [],
          "registryKeyStates": [],
          "securityResources": [],
          "triggers": [],
          "userStates": [
            {
              "aadUserId": "b9e36c12-4683-4da9-bf7d-d14f73e7bd2c",
              "accountName": "jondoe",
              "domainName": "contoso.onmicrosoft.com",
              "emailRole": "unknown",
              "isVpn": null,
              "logonDateTime": null,
              "logonId": null,
              "logonIp": null,
              "logonLocation": null,
              "logonType": null,
              "onPremisesSecurityIdentifier": null,
              "riskScore": null,
              "userAccountType": null,
              "userPrincipalName": "jondoe@contoso.onmicrosoft.com"
            }
          ],
          "uriClickSecurityStates": [],
          "vulnerabilityStates": []
        },
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
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SECURITY_ALERTS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'severity']);
  });

  it('correctly returns list of security alerts for vendor with name Azure Security Center', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts?$filter=vendorInformation/provider eq 'ASC'`) {
        return Promise.resolve(
          {
            value: alertASC
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      vendor: 'Azure Security Center'
    };

    command.action(logger, { options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(alertASC));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly returns list of security alerts for vendor with name Microsoft Cloud App Security', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts?$filter=vendorInformation/provider eq 'MCAS'`) {
        return Promise.resolve(
          {
            value: alertMCAS
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      vendor: 'Microsoft Cloud App Security'
    };

    command.action(logger, { options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(alertMCAS));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly returns list of security alerts for vendor with name Azure Active Directory Identity Protection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts?$filter=vendorInformation/provider eq 'IPC'`) {
        return Promise.resolve(
          {
            value: alertIPC
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      vendor: 'Azure Active Directory Identity Protection'
    };

    command.action(logger, { options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(alertIPC));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly returns list of security alerts as csv', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts`) {
        return Promise.resolve(
          {
            value: alertResponseCSV
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      output: "csv"
    };

    command.action(logger, { options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(alertResponseCSV));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly returns list with security alerts', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts`) {
        return Promise.resolve(
          {
            value: alertResponse
          }
        );
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
    };

    command.action(logger, { options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(alertResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when serviceAnnouncement endpoint fails', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/alerts`) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {};

    command.action(logger, { options } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Error fetching security alerts");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
