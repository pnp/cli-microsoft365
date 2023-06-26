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
const command: Command = require('./app-permission-remove');

describe(commands.APP_PERMISSION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  const validEnvironmentName = 'Default-6a2903af-9c03-4c02-a50b-e7419599925b';
  const validAppName = '784670e6-199a-4993-ae13-4b6747a0cd5d';
  const validUserId = 'd2481133-e3ed-4add-836d-6e200969dd03';
  const validUserName = 'john.doe@contoso.com';
  const validGroupId = 'c6c4b4e0-cd72-4d64-8ec2-cfbd0388ec16';
  const validGroupName = 'CLI Group';
  const appPermissionRemoveResponse = { put: [] };
  const groupResponse = {
    "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
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
    if (!auth.service.accessTokens[auth.defaultResource]) {
      auth.service.accessTokens[auth.defaultResource] = {
        expiresOn: '123',
        accessToken: 'abc'
      };
    }
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

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_PERMISSION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if appName is not GUID', async () => {
    const actual = await command.validate({ options: { appName: 'invalid', userId: validUserId, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if appName is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: validUserId, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userId is not GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: 'invalid', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, userId: validUserId, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if groupId is not GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, groupId: 'invalid', confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if groupId is a valid GUID', async () => {
    const actual = await command.validate({ options: { appName: validAppName, groupId: validGroupId, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if environmentName is used without asAdmin', async () => {
    const actual = await command.validate({ options: { appName: validAppName, environmentName: validEnvironmentName, userId: validUserId, confirm: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if environmentName is used with asAdmin', async () => {
    const actual = await command.validate({ options: { appName: validAppName, environmentName: validEnvironmentName, userId: validUserId, asAdmin: true, confirm: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the permissions for the Power App when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        appName: validAppName,
        userId: validUserId
      }
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the permissions for the Power App when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, {
      options: {
        appName: validAppName,
        userId: validUserId
      }
    });
    assert(postSpy.notCalled);
  });

  it('removes the permissions for the Power App with the user id (debug)', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, appName: validAppName, userId: validUserId, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the user id and asks for confirmation', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => {
      return { continue: true };
    });

    await command.action(logger, { options: { appName: validAppName, userId: validUserId } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the group id', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: validAppName, groupId: validGroupId, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the tenant id', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { roleName: 'CanView', appName: validAppName, tenant: true, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the username', async () => {
    sinon.stub(aadUser, 'getUserIdByUpn').resolves(validUserId);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: validAppName, userName: validUserName, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the groupname', async () => {
    sinon.stub(aadGroup, 'getGroupByDisplayName').resolves(groupResponse);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: validAppName, groupName: validGroupName, confirm: true } });
    assert(postSpy.called);
  });

  it('removes the permissions for the Power App with the user id and as admin', async () => {
    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/${validEnvironmentName}/apps/${validAppName}/modifyPermissions?api-version=2022-11-01`) {
        return appPermissionRemoveResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { appName: validAppName, userId: validUserId, environmentName: validEnvironmentName, asAdmin: true, confirm: true } });
    assert(postSpy.called);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: `Can't remove the Power App Permission`
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { appName: validAppName, userId: validUserId, confirm: true } } as any),
      new CommandError(`Can't remove the Power App Permission`));
  });
});
