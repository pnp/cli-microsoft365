import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import auth from '../../../../Auth.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-reconsent.js';
import request from '../../../../request.js';
import config from '../../../../config.js';
import { browserUtil } from '../../../../utils/browserUtil.js';
import { settingsNames } from '../../../../settingsNames.js';
import { CommandError } from '../../../../Command.js';

describe(commands.APP_RECONSENT, () => {
  const graphUrl = 'https://graph.microsoft.com/v1.0';
  const appId = '00000000-0000-0000-0000-000000000000';
  const appClientId = '11111111-1111-1111-1111-111111111111';
  const tenantId = '22222222-2222-2222-2222-222222222222';

  const servicePrincipalsResponse = {
    value: [
      {
        displayName: 'Microsoft Graph',
        appId: '00000003-0000-0000-c000-000000000000',
        servicePrincipalNames: [
          'https://dod-graph.microsoft.us',
          'https://graph.microsoft.com/',
          'https://graph.microsoft.us',
          'https://ags.windows.net',
          'https://graph.microsoft.com',
          'https://canary.graph.microsoft.com',
          '00000003-0000-0000-c000-000000000000',
          '00000003-0000-0000-c000-000000000000/ags.windows.net',
          'https://dod-graph.microsoft.us/',
          'https://graph.microsoft.us/',
          'https://canary.graph.microsoft.com'
        ],
        oauth2PermissionScopes: [
          {
            id: 'ebfcd32b-babb-40f4-a14b-42706e83bd28',
            value: 'AppCatalog.ReadWrite.All'
          },
          {
            id: 'bdfbf15f-ee85-4955-8675-146e8e5296b5',
            value: 'Application.ReadWrite.All'
          }
        ]
      },
      {
        displayName: 'Office 365 SharePoint Online',
        appId: '00000003-0000-0ff1-ce00-000000000000',
        servicePrincipalNames: [
          'https://onedrive.cloud.microsoft/',
          'https://microsoft.sharepoint-df.com',
          '00000003-0000-0ff1-ce00-000000000000',
          '00000003-0000-0ff1-ce00-000000000000/*.sharepoint.com'
        ],
        oauth2PermissionScopes: [
          {
            id: '43d8829a-ff33-456e-93cf-a7464cfa9486',
            value: 'AllSites.FullControl'
          },
          {
            id: 'aeba8e7d-0cf0-4547-9539-e49926934f39',
            value: 'TermStore.ReadWrite.All'
          }
        ]
      },
      {
        displayName: 'PowerApps Service',
        appId: '475226c6-020e-4fb2-8a90-7a972cbfc1d4',
        servicePrincipalNames: [
          'https://api.powerapps.com/',
          'https://service.powerapps.com/',
          'https://api.powerapps.com',
          'https://service.powerapps.com',
          '475226c6-020e-4fb2-8a90-7a972cbfc1d4'
        ],
        oauth2PermissionScopes: [
          {
            id: '0eb56b90-a7b5-43b5-9402-8137a8083e90',
            value: 'User'
          }
        ]
      }
    ]
  };

  let log: any[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let patchStub: sinon.SinonStub;
  let browserStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.appId = appClientId;
    auth.connection.tenant = tenantId;

    browserStub = sinon.stub(browserUtil, 'open').resolves();
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
    loggerSpy = sinon.spy(logger, 'log');

    sinon.stub(config, 'allScopes').value([
      'https://graph.microsoft.com/AppCatalog.ReadWrite.All',
      'https://graph.microsoft.com/Application.ReadWrite.All',
      'https://microsoft.sharepoint-df.com/AllSites.FullControl',
      'https://microsoft.sharepoint-df.com/TermStore.ReadWrite.All',
      'https://api.powerapps.com//User'
    ]);

    patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications/${appId}`) {
        return;
      }

      throw 'Invalid request with URL: ' + opts.url;
    });

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((name: string, defaultValue: any) => defaultValue);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      config.allScopes,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.deactivate();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_RECONSENT);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly logs output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications?$filter=appId eq '${appClientId}'&$select=requiredResourceAccess,id`) {
        return {
          value: [
            {
              id: appId,
              requiredResourceAccess: []
            }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals?$select=displayName,appId,oauth2PermissionScopes,servicePrincipalNames`) {
        return servicePrincipalsResponse;
      }

      throw 'Invalid request with URL: ' + opts.url;
    });

    await command.action(logger, { options: {} });
    assert(loggerSpy.calledOnceWith(`To consent to the new scopes for your Microsoft Entra application registration, please navigate to the following URL: https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${appClientId}`));
  });

  it('correctly adds new scopes to the app registration', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications?$filter=appId eq '${appClientId}'&$select=requiredResourceAccess,id`) {
        return {
          value: [
            {
              id: appId,
              requiredResourceAccess: [
                {
                  resourceAppId: '00000003-0000-0000-c000-000000000000',
                  resourceAccess: [
                    {
                      id: 'bdfbf15f-ee85-4955-8675-146e8e5296b5',
                      type: 'Scope'
                    }
                  ]
                }
              ]
            }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals?$select=displayName,appId,oauth2PermissionScopes,servicePrincipalNames`) {
        return servicePrincipalsResponse;
      }

      throw 'Invalid GET request with URL: ' + opts.url;
    });

    await command.action(logger, { options: { verbose: true } });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      requiredResourceAccess: [
        {
          resourceAppId: '00000003-0000-0000-c000-000000000000',
          resourceAccess: [
            {
              id: 'bdfbf15f-ee85-4955-8675-146e8e5296b5',
              type: 'Scope'
            },
            {
              id: 'ebfcd32b-babb-40f4-a14b-42706e83bd28',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '00000003-0000-0ff1-ce00-000000000000',
          resourceAccess: [
            {
              id: '43d8829a-ff33-456e-93cf-a7464cfa9486',
              type: 'Scope'
            },
            {
              id: 'aeba8e7d-0cf0-4547-9539-e49926934f39',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '475226c6-020e-4fb2-8a90-7a972cbfc1d4',
          resourceAccess: [
            {
              id: '0eb56b90-a7b5-43b5-9402-8137a8083e90',
              type: 'Scope'
            }
          ]
        }
      ]
    });
  });

  it('correctly adds new scopes and does not remove existing ones', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications?$filter=appId eq '${appClientId}'&$select=requiredResourceAccess,id`) {
        return {
          value: [
            {
              id: appId,
              requiredResourceAccess: [
                {
                  resourceAppId: '00000003-0000-0000-c000-000000000000',
                  resourceAccess: [
                    {
                      id: 'ebf0f66e-9fb1-49e4-a278-222f76911cf4',
                      type: 'Scope'
                    }
                  ]
                }
              ]
            }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals?$select=displayName,appId,oauth2PermissionScopes,servicePrincipalNames`) {
        return servicePrincipalsResponse;
      }

      throw 'Invalid GET request with URL: ' + opts.url;
    });

    await command.action(logger, { options: { verbose: true } });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      requiredResourceAccess: [
        {
          resourceAppId: '00000003-0000-0000-c000-000000000000',
          resourceAccess: [
            {
              id: 'ebf0f66e-9fb1-49e4-a278-222f76911cf4',
              type: 'Scope'
            },
            {
              id: 'ebfcd32b-babb-40f4-a14b-42706e83bd28',
              type: 'Scope'
            },
            {
              id: 'bdfbf15f-ee85-4955-8675-146e8e5296b5',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '00000003-0000-0ff1-ce00-000000000000',
          resourceAccess: [
            {
              id: '43d8829a-ff33-456e-93cf-a7464cfa9486',
              type: 'Scope'
            },
            {
              id: 'aeba8e7d-0cf0-4547-9539-e49926934f39',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '475226c6-020e-4fb2-8a90-7a972cbfc1d4',
          resourceAccess: [
            {
              id: '0eb56b90-a7b5-43b5-9402-8137a8083e90',
              type: 'Scope'
            }
          ]
        }
      ]
    });
  });

  it('does not fail when an unknown principal or scope is defined', async () => {
    const invalidScopes = [...config.allScopes];
    invalidScopes.push('https://invalid.microsoft.com/ChannelMessage.Send', 'https://graph.microsoft.com/invalidScope');
    sinonUtil.restore(config.allScopes);
    sinon.stub(config, 'allScopes').value(invalidScopes);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications?$filter=appId eq '${appClientId}'&$select=requiredResourceAccess,id`) {
        return {
          value: [
            {
              id: appId,
              requiredResourceAccess: []
            }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals?$select=displayName,appId,oauth2PermissionScopes,servicePrincipalNames`) {
        return servicePrincipalsResponse;
      }

      throw 'Invalid request with URL: ' + opts.url;
    });

    await command.action(logger, { options: { verbose: true } });
    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      requiredResourceAccess: [
        {
          resourceAppId: '00000003-0000-0000-c000-000000000000',
          resourceAccess: [
            {
              id: 'ebfcd32b-babb-40f4-a14b-42706e83bd28',
              type: 'Scope'
            },
            {
              id: 'bdfbf15f-ee85-4955-8675-146e8e5296b5',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '00000003-0000-0ff1-ce00-000000000000',
          resourceAccess: [
            {
              id: '43d8829a-ff33-456e-93cf-a7464cfa9486',
              type: 'Scope'
            },
            {
              id: 'aeba8e7d-0cf0-4547-9539-e49926934f39',
              type: 'Scope'
            }
          ]
        },
        {
          resourceAppId: '475226c6-020e-4fb2-8a90-7a972cbfc1d4',
          resourceAccess: [
            {
              id: '0eb56b90-a7b5-43b5-9402-8137a8083e90',
              type: 'Scope'
            }
          ]
        }
      ]
    });
  });

  it('automatically opens URL in browser if setting is active', async () => {
    sinonUtil.restore(cli.getSettingWithDefaultValue);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((name: string, defaultValue: any) => name === settingsNames.autoOpenLinksInBrowser ? true : defaultValue);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${graphUrl}/applications?$filter=appId eq '${appClientId}'&$select=requiredResourceAccess,id`) {
        return {
          value: [
            {
              id: appId,
              requiredResourceAccess: []
            }
          ]
        };
      }

      if (opts.url === `${graphUrl}/servicePrincipals?$select=displayName,appId,oauth2PermissionScopes,servicePrincipalNames`) {
        return servicePrincipalsResponse;
      }

      throw 'Invalid request with URL: ' + opts.url;
    });

    await command.action(logger, { options: {} });
    assert(browserStub.calledOnceWith(`https://login.microsoftonline.com/${tenantId}/adminconsent?client_id=${appClientId}`));
  });

  it('correctly handles error when appId is not found', async () => {
    sinon.stub(request, 'get').rejects({ error: { message: 'Insufficient privileges to complete the operation.' } });

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError('Insufficient privileges to complete the operation.'));
  });
});
