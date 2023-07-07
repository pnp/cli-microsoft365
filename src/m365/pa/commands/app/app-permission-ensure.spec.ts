import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { aadUser } from '../../../../utils/aadUser';
import { aadGroup } from '../../../../utils/aadGroup';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./app-permission-ensure');

describe(commands.APP_PERMISSION_ENSURE, () => {
  let cli: Cli;
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
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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

  it('passes validation if tenant is used wit role CanView', async () => {
    const actual = await command.validate({ options: { roleName: 'CanView', appName: validAppName, tenant: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('updates permissions to a Power App by userId and sends invitation mail', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'Notify'
          && opts.data.put[0].properties.principal.id === validUserId
          && opts.data.put[0].properties.principal.type === 'User'
          && opts.data.put[0].properties.roleName === validRoleName) {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, userId: validUserId, sendInvitationMail: true } });
    assert(postSpy.called);
  });

  it('updates permissions to a Power App with by groupId', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'DoNotNotify'
          && opts.data.put[0].properties.principal.id === validGroupId
          && opts.data.put[0].properties.principal.type === 'Group'
          && opts.data.put[0].properties.roleName === validRoleName) {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, groupId: validGroupId } });
    assert(postSpy.called);
  });

  it('shares a Power App with the entire tenant', async () => {
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: '123',
      accessToken: 'abc'
    };
    sinon.stub(accessToken, 'getTenantIdFromAccessToken').resolves(tenantId);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'DoNotNotify'
          && opts.data.put[0].properties.principal.id === tenantId
          && opts.data.put[0].properties.principal.type === 'Tenant'
          && opts.data.put[0].properties.roleName === 'CanView') {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roleName: 'CanView', appName: validAppName, tenant: true } });
    assert(postSpy.called);
  });

  it('updates permissions to a Power App by userName', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').resolves(validUserId);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'DoNotNotify'
          && opts.data.put[0].properties.principal.id === validUserId
          && opts.data.put[0].properties.principal.type === 'User'
          && opts.data.put[0].properties.roleName === validRoleName) {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, userName: validUserName } });
    assert(postSpy.called);
  });

  it('updates permissions to a Power App by groupName', async () => {
    sinon.stub(aadGroup, 'getGroupByDisplayName').resolves(groupResponse);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'DoNotNotify'
          && opts.data.put[0].properties.principal.id === validGroupId
          && opts.data.put[0].properties.principal.type === 'Group'
          && opts.data.put[0].properties.roleName === validRoleName) {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, roleName: validRoleName, appName: validAppName, groupName: validGroupName } });
    assert(postSpy.called);
  });

  it('updates permissions to a Power App with by userId as admin', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        if (opts.data.put[0].properties.NotifyShareTargetOption === 'DoNotNotify'
          && opts.data.put[0].properties.principal.id === validUserId
          && opts.data.put[0].properties.principal.type === 'User'
          && opts.data.put[0].properties.roleName === validRoleName) {
          return appPermissionEnsureResponse;
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { roleName: validRoleName, appName: validAppName, userId: validUserId, environmentName: validEnvironmentName, asAdmin: true } });
    assert(postSpy.called);
  });

  it('correctly handles API OData error', async () => {
    const errorMessage = 'Can\'t update the Power App Permission';
    sinon.stub(request, 'post').rejects({
      error: {
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { roleName: validRoleName, appName: validAppName, userId: validUserId } } as any),
      new CommandError(errorMessage));
  });
});
