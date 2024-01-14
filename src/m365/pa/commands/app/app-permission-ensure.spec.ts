import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraGroup } from '../../../../utils/entraGroup.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './app-permission-ensure.js';

describe(commands.APP_PERMISSION_ENSURE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const validEnvironmentName = 'Default-6a2903af-9c03-4c02-a50b-e7419599925b';
  const validAppName = '784670e6-199a-4993-ae13-4b6747a0cd5d';
  const validUserId = 'd2481133-e3ed-4add-836d-6e200969dd03';
  const validUserName = 'john.doe@contoso.com';
  const validGroupId = 'c6c4b4e0-cd72-4d64-8ec2-cfbd0388ec16';
  const validGroupName = 'CLI Group';
  const validRoleName = 'CanEdit';
  const tenantId = '174290ec-373f-4d4c-89ea-9801dad0acd9';
  const appPermissionEnsureResponse = {
    put: [
      {
        id: "/providers/Microsoft.PowerApps/scopes/admin/apps/784670e6-199a-4993-ae13-4b6747a0cd5/permissions/241adbf6-2a56-4c72-81f2-69e75de6ac34",
        properties: {
          roleName: "CanView",
          scope: "/providers/Microsoft.PowerApps/apps/784670e6-199a-4993-ae13-4b6747a0cd5",
          principal: {
            id: "d2481133-e3ed-4add-836d-6e200969dd03",
            type: "User"
          },
          resourceResponses: [
            {
              id: "/providers/Microsoft.PowerApps/apps/784670e6-199a-4993-ae13-4b6747a0cd5",
              statusCode: "Created",
              responseCode: "ResourceShared",
              message: "This was shared with 'CanView' permission.",
              type: "/providers/Microsoft.PowerApps/apps"
            }
          ]
        }
      }
    ]
  };
  const groupResponse = {
    "id": validGroupId,
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2017-11-29T03:27:05Z",
    "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
    "displayName": "Finance",
    "groupTypes": [
      "Unified"
    ],
    "mail": "finance@contoso.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "finance",
    "onPremisesLastSyncDateTime": null,
    "onPremisesProvisioningErrors": [],
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "proxyAddresses": [
      "SMTP:finance@contoso.onmicrosoft.com"
    ],
    "renewedDateTime": "2017-11-29T03:27:05Z",
    "securityEnabled": false,
    "visibility": "Public"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((_, defaultValue) => defaultValue);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      accessToken.getTenantIdFromAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_ENSURE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appName is not a GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: 'invalid', userId: validUserId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appName is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not a GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a UPN', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userName: 'John Doe' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userName is a valid UPN', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userName: validUserName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if groupId is not a GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, groupId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, groupId: validGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if roleName is not a valid role', async () => {
    const actual = await command.validate({ options: { roleName: 'invalid', appName: validAppName, userId: validUserId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if roleName is a valid role', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userId: validUserId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if asAdmin is used without environmentName', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, userId: validUserId, asAdmin: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if environmentName is used without asAdmin', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, environmentName: validEnvironmentName, userId: validUserId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if environmentName is used with asAdmin', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, environmentName: validEnvironmentName, userId: validUserId, asAdmin: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if tenant is used without role CanView', async () => {
    const actual = await command.validate({ options: { roleName: validRoleName, appName: validAppName, tenant: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if tenant is used with role CanView', async () => {
    const actual = await command.validate({ options: { roleName: 'CanView', appName: validAppName, tenant: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('updates permissions to a Power App by userId and sends invitation mail', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId,
              type: 'User'
            },
            NotifyShareTargetOption: 'Notify',
            roleName: validRoleName
          }
        }
      ]
    };

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, userId: validUserId, sendInvitationMail: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('updates permissions to a Power App with by groupId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validGroupId,
              type: 'Group'
            },
            NotifyShareTargetOption: 'DoNotNotify',
            roleName: validRoleName
          }
        }
      ]
    };

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, groupId: validGroupId } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('shares a Power App with the entire tenant', async () => {
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').resolves(tenantId);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: tenantId,
              type: 'Tenant'
            },
            NotifyShareTargetOption: 'DoNotNotify',
            roleName: 'CanView'
          }
        }
      ]
    };

    await command.action(logger, { options: { verbose: true, roleName: 'CanView', appName: validAppName, tenant: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('updates permissions to a Power App by userName', async () => {
    sinon.stub(entraUser, 'getUserIdByUpn').resolves(validUserId);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId,
              type: 'User'
            },
            NotifyShareTargetOption: 'DoNotNotify',
            roleName: validRoleName
          }
        }
      ]
    };

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, userName: validUserName } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('updates permissions to a Power App by groupName', async () => {
    sinon.stub(entraGroup, 'getGroupByDisplayName').resolves(groupResponse);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validGroupId,
              type: 'Group'
            },
            NotifyShareTargetOption: 'DoNotNotify',
            roleName: validRoleName
          }
        }
      ]
    };

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, groupName: validGroupName } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('updates permissions to a Power App with by userId as admin', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionEnsureResponse;
      }

      throw 'Invalid request';
    });

    const requestBody = {
      put: [
        {
          properties: {
            principal: {
              id: validUserId,
              type: 'User'
            },
            NotifyShareTargetOption: 'DoNotNotify',
            roleName: validRoleName
          }
        }
      ]
    };

    await command.action(logger, { options: { roleName: validRoleName, appName: validAppName, userId: validUserId, environmentName: validEnvironmentName, asAdmin: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, requestBody);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Can\'t update the Power App Permission';
    sinon.stub(request, 'post').rejects({
      error: {
        error: {
          message: errorMessage
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { roleName: validRoleName, appName: validAppName, userId: validUserId } } as any),
      new CommandError(errorMessage));
  });
});
