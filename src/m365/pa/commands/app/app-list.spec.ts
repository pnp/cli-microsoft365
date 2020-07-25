import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./app-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.APP_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves apps (debug)', (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": apps });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '4d4bb961-eef9-4258-8516-aa8d64e6b477',
            displayName: 'Request-a-team'
          },
          {
            name: '79506a60-9c4c-4798-a1fa-aea552ef046e',
            displayName: 'Toolkit'
          },
          {
            name: 'f581c872-9852-4100-8e25-3d6891595204',
            displayName: 'Playwright'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves apps', (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": apps });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            name: '4d4bb961-eef9-4258-8516-aa8d64e6b477',
            displayName: 'Request-a-team'
          },
          {
            name: '79506a60-9c4c-4798-a1fa-aea552ef046e',
            displayName: 'Toolkit'
          },
          {
            name: 'f581c872-9852-4100-8e25-3d6891595204',
            displayName: 'Playwright'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all properties when output is JSON', (done) => {
    const apps = [
      {
        "name": "4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "id": "/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "1366",
          "primaryDeviceHeight": "768",
          "sienaVersion": "20200512T062535Z-3.20023.8.0",
          "deviceCapabilities": "",
          "supportsPortrait": "false",
          "supportsLandscape": "true",
          "primaryFormFactor": "Tablet",
          "publisherVersion": "3.20023.8",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false"
        },
        "properties": {
          "appVersion": "2020-07-08T12:28:37Z",
          "lastDraftVersion": "2020-07-08T12:28:37Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20023,
            "build": 8,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(37, 62, 143, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Request-a-team",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-07-08T12:28:37.957179Z",
          "lastModifiedTime": "2020-07-08T12:28:38.7556554Z",
          "lastPublishTime": "2020-07-08T12:28:37Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477",
          "appOpenUri": "https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "a65df3f8-e66c-4cbd-b13f-458b7e96f677": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365groups",
              "displayName": "Office 365 Groups",
              "iconUri": "https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png",
              "dataSources": [
                "Office365Groups"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "041cbeda-55ca-4c48-b8e3-03928fb72bb2": {
              "id": "/providers/microsoft.powerapps/apis/shared_logicflows",
              "displayName": "Logic flows",
              "iconUri": "https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png",
              "dataSources": [
                "CheckTeamAvailability"
              ],
              "dependencies": [
                "97e5ce6b-9f9a-4186-885f-9b5d6476c732"
              ],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "97e5ce6b-9f9a-4186-885f-9b5d6476c732": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [],
              "dependencies": [],
              "dependents": [
                "041cbeda-55ca-4c48-b8e3-03928fb72bb2"
              ],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            },
            "00deca03-387b-4ad4-bbd4-cefc640d1c9b": {
              "id": "/providers/microsoft.powerapps/apis/shared_sharepointonline",
              "displayName": "SharePoint",
              "iconUri": "https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png",
              "dataSources": [
                "Teams Templates",
                "Teams Requests",
                "Team Request Settings"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {
                "https://m365x023142.sharepoint.com/sites/RequestateamApp": {
                  "dataSources": {
                    "Teams Templates": {
                      "tableName": "298485ad-73cc-4b5f-a013-b56111ec351a"
                    },
                    "Teams Requests": {
                      "tableName": "a471ecf0-01f3-4e3e-902b-b48daaa23aba"
                    },
                    "Team Request Settings": {
                      "tableName": "3770cede-bff2-42a6-ba12-2f4cbccb90d3"
                    }
                  }
                }
              },
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": false,
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "79506a60-9c4c-4798-a1fa-aea552ef046e",
        "id": "/providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "minimumRequiredApiVersion": "2.2.0"
        },
        "properties": {
          "appVersion": "2020-06-08T20:52:24Z",
          "lastDraftVersion": "2020-06-08T20:52:24Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 18114,
            "build": 26,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N90efe94c-af45-4639-885e-d69f32cd6c9f-logoSmallFile?sv=2018-03-28&sr=c&sig=mm7Cj0z%2FlX42FaSCSA9MtwBxMVEEnveqb1%2FsQhfLsRw%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Toolkit",
          "description": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200608t000000z1cbf48a3f3b54583b8932510cbdf20b0/6342866521103212774/N1bc8ee4e-1a31-4917-86f8-b9309667d09b-document.msapp?sv=2018-03-28&sr=c&sig=vcph4RCCqlB6Hc78oScTcfdkMfj6dMggsvPxxqrBVpU%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-06-08T20:52:24.1796831Z",
          "lastModifiedTime": "2020-06-08T20:52:24.4140538Z",
          "lastPublishTime": "2020-06-08T20:52:24Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/79506a60-9c4c-4798-a1fa-aea552ef046e",
          "appOpenUri": "https://apps.powerapps.com/play/79506a60-9c4c-4798-a1fa-aea552ef046e?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "databaseReferences": {},
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "almMode": "Environment",
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "appType": "ClassicCanvasApp"
      },
      {
        "name": "f581c872-9852-4100-8e25-3d6891595204",
        "id": "/providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
        "type": "Microsoft.PowerApps/apps",
        "tags": {
          "primaryDeviceWidth": "640",
          "primaryDeviceHeight": "1136",
          "sienaVersion": "20200812T204016Z-3.20074.20.0",
          "deviceCapabilities": "",
          "supportsPortrait": "true",
          "supportsLandscape": "false",
          "primaryFormFactor": "Phone",
          "publisherVersion": "3.20074.20",
          "minimumRequiredApiVersion": "2.2.0",
          "hasComponent": "false",
          "hasUnlockedComponent": "false",
          "isUnifiedRootApp": "false"
        },
        "properties": {
          "appVersion": "2020-08-12T20:40:16Z",
          "lastDraftVersion": "2020-08-12T20:40:16Z",
          "lifeCycleId": "Published",
          "status": "Ready",
          "createdByClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "minClientVersion": {
            "major": 3,
            "minor": 20074,
            "build": 20,
            "revision": 0,
            "majorRevision": 0,
            "minorRevision": 0
          },
          "owner": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "createdBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastModifiedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "lastPublishedBy": {
            "id": "a86f34fb-fc0b-476f-b2d3-84b2648cc87a",
            "displayName": "Garry Trinder",
            "email": "garry@trinder365dev.onmicrosoft.com",
            "type": "User",
            "tenantId": "e8954f17-a373-4b61-b54d-45c038fe3188",
            "userPrincipalName": "garry@trinder365dev.onmicrosoft.com"
          },
          "backgroundColor": "rgba(0, 176, 240, 1)",
          "backgroundImageUri": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/logoSmallFile?sv=2018-03-28&sr=c&sig=sqK6%2FXY4cHidwE%2Brb3JoBV3bNToOaA6EM3%2FczbWMQDc%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsColorIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsColorIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "teamsOutlineIconUrl": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/teamsOutlineIcon.png?sv=2018-03-28&sr=c&sig=UYs6LV%2BGqPjfNczXP80lm%2BmG1ebFNcLCF0D8MIJ6Lt8%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl",
          "displayName": "Playwright",
          "description": "",
          "commitMessage": "",
          "appUris": {
            "documentUri": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z1766ec3fd78941bea695c957e898a62a/document.msapp?sv=2018-03-28&sr=c&sig=aToV3yl8gK0eiAPsh3DIxo3VC77OyLrZgYo2G%2BYXDgI%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "imageUris": []
          },
          "createdTime": "2020-08-10T23:28:41.8191546Z",
          "lastModifiedTime": "2020-08-12T20:40:20.3706202Z",
          "lastPublishTime": "2020-08-12T20:40:20.3706202Z",
          "sharedGroupsCount": 0,
          "sharedUsersCount": 0,
          "appOpenProtocolUri": "ms-apps:///providers/Microsoft.PowerApps/apps/f581c872-9852-4100-8e25-3d6891595204",
          "appOpenUri": "https://apps.powerapps.com/play/f581c872-9852-4100-8e25-3d6891595204?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188",
          "connectionReferences": {
            "dd1ebcc1-9930-4e87-a680-45fb1eaf94e6": {
              "id": "/providers/microsoft.powerapps/apis/shared_office365users",
              "displayName": "Office 365 Users",
              "iconUri": "https://connectoricons-prod.azureedge.net/releases/v1.0.1381/1.0.1381.2096/office365users/icon.png",
              "dataSources": [
                "Office365Users"
              ],
              "dependencies": [],
              "dependents": [],
              "isOnPremiseConnection": false,
              "bypassConsent": false,
              "dataSets": {},
              "apiTier": "Standard",
              "isCustomApiConnection": false
            }
          },
          "userAppMetadata": {
            "favorite": "NotSpecified",
            "lastOpenedTime": "2020-08-13T23:26:44.2982102Z",
            "includeInAppsList": true
          },
          "isFeaturedApp": false,
          "bypassConsent": false,
          "isHeroApp": false,
          "environment": {
            "id": "/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188",
            "name": "Default-e8954f17-a373-4b61-b54d-45c038fe3188"
          },
          "appPackageDetails": {
            "playerPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/player.msappk?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "sizeInBytes": 0
            },
            "webPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net:443/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl",
              "readonlyValue": "https://pafeblobprodln-secondary.blob.core.windows.net/20200812t000000z4928635e44124aa5b50bfae36ed252b5/05358dc3-4770-4f73-a5ec-fe0a3e341454/web/index.web.html?sv=2018-03-28&sr=c&sig=UXTet030wmU8QR2TH8TWCrgm354F2LTjgIubPcfXGD4%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl"
            },
            "unauthenticatedWebPackage": {
              "value": "https://pafeblobprodln.blob.core.windows.net/alt20200810t000000zc57cd52652b24a1eb573f7b2a36a10a9/20200812T204028Z/index.web.html"
            },
            "documentServerVersion": {
              "major": 3,
              "minor": 20074,
              "build": 20,
              "revision": 0,
              "majorRevision": 0,
              "minorRevision": 0
            },
            "appPackageResourcesKind": "Split",
            "packagePropertiesJson": "{\"cdnUrl\":\"https://content.powerapps.com/resource/app\",\"preLoadIdx\":\"https://content.powerapps.com/resource/app/4g3nunecadgk9/preloadindex.web.html\",\"id\":\"637328616254057865\",\"v\":2.1}"
          },
          "almMode": "Environment",
          "performanceOptimizationEnabled": true,
          "unauthenticatedWebPackageHint": "1eef5df9-6032-459c-9194-77d926b11f37",
          "canConsumeAppPass": true,
          "appPlanClassification": "Standard",
          "usesPremiumApi": false,
          "usesOnlyGrandfatheredPremiumApis": true,
          "usesCustomApi": false,
          "usesOnPremiseGateway": false
        },
        "isAppComponentLibrary": false,
        "appType": "ClassicCanvasApp"
      }
    ];

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`providers/Microsoft.PowerApps/apps?api-version=2017-08-01`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "value": apps });
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(apps));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps found (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.resolve({ value: [] });
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('No apps found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        error: {
          'odata.error': {
            code: '-1, InvalidOperationException',
            message: {
              value: 'An error has occurred'
            }
          }
        }
      });
    });

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});