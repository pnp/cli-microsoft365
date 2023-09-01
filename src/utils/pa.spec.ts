import * as assert from 'assert';
import * as sinon from 'sinon';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { pa } from './pa.js';
import { Logger } from '../cli/Logger.js';

const validDisplayName = 'Request-a-team';
const appResponse = {
  name: '4d4bb961-eef9-4258-8516-aa8d64e6b477',
  id: '/providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477',
  type: 'Microsoft.PowerApps/apps',
  tags: {
    primaryDeviceWidth: 1366,
    primaryDeviceHeight: 768,
    sienaVersion: '20200512T062535Z-3.20023.8.0',
    deviceCapabilities: '',
    supportsPortrait: false,
    supportsLandscape: true,
    primaryFormFactor: 'Tablet',
    publisherVersion: '3.20023.8',
    minimumRequiredApiVersion: '2.2.0',
    hasComponent: false,
    hasUnlockedComponent: false
  },
  properties: {
    appVersion: '2020-07-08T12:28:37Z',
    lastDraftVersion: '2020-07-08T12:28:37Z',
    lifeCycleId: 'Published',
    status: 'Ready',
    createdByClientVersion: {
      major: 3,
      minor: 20023,
      build: 8,
      revision: 0,
      majorRevision: 0,
      minorRevision: 0
    },
    minClientVersion: {
      major: 3,
      minor: 20023,
      build: 8,
      revision: 0,
      majorRevision: 0,
      minorRevision: 0
    },
    owner: {
      id: 'a86f34fb-fc0b-476f-b2d3-84b2648cc87a',
      displayName: 'John Doe',
      email: 'john.doe@contoso.onmicrosoft.com',
      type: 'User',
      tenantId: 'e8954f17-a373-4b61-b54d-45c038fe3188',
      userPrincipalName: 'john.doe@contoso.onmicrosoft.com'
    },
    createdBy: {
      id: 'a86f34fb-fc0b-476f-b2d3-84b2648cc87a',
      displayName: 'John Doe',
      email: 'john.doe@contoso.onmicrosoft.com',
      type: 'User',
      tenantId: 'e8954f17-a373-4b61-b54d-45c038fe3188',
      userPrincipalName: 'john.doe@contoso.onmicrosoft.com'
    },
    lastModifiedBy: {
      id: 'a86f34fb-fc0b-476f-b2d3-84b2648cc87a',
      displayName: 'John Doe',
      email: 'john.doe@contoso.onmicrosoft.com',
      type: 'User',
      tenantId: 'e8954f17-a373-4b61-b54d-45c038fe3188',
      userPrincipalName: 'john.doe@contoso.onmicrosoft.com'
    },
    lastPublishedBy: {
      id: 'a86f34fb-fc0b-476f-b2d3-84b2648cc87a',
      displayName: 'John Doe',
      email: 'john.doe@contoso.onmicrosoft.com',
      type: 'User',
      tenantId: 'e8954f17-a373-4b61-b54d-45c038fe3188',
      userPrincipalName: 'john.doe@contoso.onmicrosoft.com'
    },
    backgroundColor: 'rgba(37, 62, 143, 1)',
    backgroundImageUri: 'https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N0eb33631-4950-45e8-b569-8ba8611af629-logoSmallFile?sv=2018-03-28&sr=c&sig=rTJyePWWDMM6mvIhZaOkRsEdLxFE4X6UGXjrqrz3iYo%3D&se=2020-10-05T18%3A56%3A46Z&sp=rl',
    displayName: 'Request-a-team',
    description: '',
    commitMessage: '',
    appUris: {
      documentUri: {
        value: 'https://pafeblobprodln.blob.core.windows.net:443/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl',
        readonlyValue: 'https://pafeblobprodln-secondary.blob.core.windows.net/20200708t000000z4d9d5509e6c745d3bbd4d6d317890ccd/13103204444004720806/N9d70c8fe-cbc0-4226-8818-372c4261e0c6-document.msapp?sv=2018-03-28&sr=c&sig=ltod6hA3brZQF9qTxNKFg0ryuX7IxsrJLY8KdA9u8f8%3D&se=2020-08-28T08%3A00%3A00Z&sp=rl'
      },
      imageUris: []
    },
    createdTime: '2020-07-08T12:28:37.957179Z',
    lastModifiedTime: '2020-07-08T12:28:38.7556554Z',
    lastPublishTime: '2020-07-08T12:28:37Z',
    sharedGroupsCount: 0,
    sharedUsersCount: 0,
    appOpenProtocolUri: 'ms-apps:///providers/Microsoft.PowerApps/apps/4d4bb961-eef9-4258-8516-aa8d64e6b477',
    appOpenUri: 'https://apps.powerapps.com/play/4d4bb961-eef9-4258-8516-aa8d64e6b477?tenantId=e8954f17-a373-4b61-b54d-45c038fe3188',
    connectionReferences: {
      '9d5036a3-8b23-4125-a5cc-7dc0dbb2f8cb': {
        id: '/providers/microsoft.powerapps/apis/shared_office365users',
        displayName: 'Office 365 Users',
        iconUri: 'https://connectoricons-prod.azureedge.net/office365users/icon_1.0.1357.2029.png',
        dataSources: [
          'Office365Users'
        ],
        dependencies: [],
        dependents: [],
        isOnPremiseConnection: false,
        bypassConsent: false,
        dataSets: {},
        apiTier: 'Standard',
        isCustomApiConnection: false
      },
      'a65df3f8-e66c-4cbd-b13f-458b7e96f677': {
        id: '/providers/microsoft.powerapps/apis/shared_office365groups',
        displayName: 'Office 365 Groups',
        iconUri: 'https://connectoricons-prod.azureedge.net/office365groups/icon_1.0.1329.1953.png',
        dataSources: [
          'Office365Groups'
        ],
        dependencies: [],
        dependents: [],
        isOnPremiseConnection: false,
        bypassConsent: false,
        dataSets: {},
        apiTier: 'Standard',
        isCustomApiConnection: false
      },
      '041cbeda-55ca-4c48-b8e3-03928fb72bb2': {
        id: '/providers/microsoft.powerapps/apis/shared_logicflows',
        displayName: 'Logic flows',
        iconUri: 'https://resourcestackdeploy.blob.core.windows.net/scripts/13276078.png',
        dataSources: [
          'CheckTeamAvailability'
        ],
        dependencies: [
          '97e5ce6b-9f9a-4186-885f-9b5d6476c732'
        ],
        dependents: [],
        isOnPremiseConnection: false,
        bypassConsent: false,
        dataSets: {},
        apiTier: 'Standard',
        isCustomApiConnection: false
      },
      '97e5ce6b-9f9a-4186-885f-9b5d6476c732': {
        id: '/providers/microsoft.powerapps/apis/shared_sharepointonline',
        displayName: 'SharePoint',
        iconUri: 'https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png',
        dataSources: [],
        dependencies: [],
        dependents: [
          '041cbeda-55ca-4c48-b8e3-03928fb72bb2'
        ],
        isOnPremiseConnection: false,
        bypassConsent: false,
        dataSets: {},
        apiTier: 'Standard',
        isCustomApiConnection: false
      },
      '00deca03-387b-4ad4-bbd4-cefc640d1c9b': {
        id: '/providers/microsoft.powerapps/apis/shared_sharepointonline',
        displayName: 'SharePoint',
        iconUri: 'https://connectoricons-prod.azureedge.net/sharepointonline/icon_1.0.1363.2042.png',
        dataSources: [
          'Teams Templates',
          'Teams Requests',
          'Team Request Settings'
        ],
        dependencies: [],
        dependents: [],
        isOnPremiseConnection: false,
        bypassConsent: false,
        dataSets: {
          'https://contoso.sharepoint.com/sites/RequestateamApp': {
            dataSources: {
              'Teams Templates': {
                tableName: '298485ad-73cc-4b5f-a013-b56111ec351a'
              },
              'Teams Requests': {
                tableName: 'a471ecf0-01f3-4e3e-902b-b48daaa23aba'
              },
              'Team Request Settings': {
                tableName: '3770cede-bff2-42a6-ba12-2f4cbccb90d3'
              }
            }
          }
        },
        apiTier: 'Standard',
        isCustomApiConnection: false
      }
    },
    databaseReferences: {},
    userAppMetadata: {
      favorite: 'NotSpecified',
      includeInAppsList: true
    },
    isFeaturedApp: false,
    bypassConsent: false,
    isHeroApp: false,
    environment: {
      id: '/providers/Microsoft.PowerApps/environments/Default-e8954f17-a373-4b61-b54d-45c038fe3188',
      name: 'Default-e8954f17-a373-4b61-b54d-45c038fe3188'
    },
    almMode: 'Environment',
    performanceOptimizationEnabled: false,
    canConsumeAppPass: true,
    appPlanClassification: 'Standard',
    usesPremiumApi: false,
    usesOnlyGrandfatheredPremiumApis: true,
    usesCustomApi: false,
    usesOnPremiseGateway: false
  },
  isAppComponentLibrary: false,
  appType: 'ClassicCanvasApp'
};

describe('utils/pa', () => {
  let logger: Logger;
  let log: string[];

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
  });

  it('correctly get a Power App by displayname.', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://api.powerapps.com/providers/Microsoft.PowerApps/apps?api-version=2017-08-01') {
        return { value: [appResponse] };
      }

      return 'Invalid Request';
    });

    const actual = await pa.getAppByDisplayName(validDisplayName, logger, true);
    assert.strictEqual(actual, appResponse);
  });

  it('throws error message when no Power App was found by displayname', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://api.powerapps.com/providers/Microsoft.PowerApps/apps?api-version=2017-08-01') {
        return { value: [appResponse] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(pa.getAppByDisplayName('another app'), `An error occured`);
  });

  it('throws error message when no Power Apps were found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://api.powerapps.com/providers/Microsoft.PowerApps/apps?api-version=2017-08-01') {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(pa.getAppByDisplayName(validDisplayName), `An error occured`);
  });
}); 