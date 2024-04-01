import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-app-upgrade.js';
import { odata } from '../../../../utils/odata.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.USER_APP_UPGRADE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const userId = '520883a5-a09f-4ff4-bfbc-4ae8b57b1e39';
  const userName = 'adelev@contoso.com';
  const installedAppId = 'ODMyM2Y3ZmUtZThhNC00NmM0LWI1ZWEtZjQ4NjQ4ODdkMTYwIyNlZmZlMTNlZi0wMGUyLTRkNTUtYTIwNy1mMmQ5ZDFkZDcyNjM=';
  const installedAppName = 'Custom App Name';

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.post,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_APP_UPGRADE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({ options: { userId: 'invalid', id: installedAppId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN.', async () => {
    const actual = await command.validate({ options: { userName: 'invalid', id: installedAppId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the userName is a valid userName', async () => {
    const actual = await command.validate({ options: { id: installedAppId, userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the userId is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: installedAppId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('upgrades application retrieved by name and handles multiple values found', async () => {
    const userInstalledAppsResponse = [{ id: installedAppId }, { id: 'installedTeamsAppId' }];

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(installedAppName)}'&$select=id`) {
        return userInstalledAppsResponse;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(userInstalledAppsResponse[0]);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps/${installedAppId}/upgrade`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: userId, name: installedAppName } });
    assert(postStub.calledOnce);
  });

  it('upgrades application retrieved by name when only single application is found', async () => {
    const userInstalledAppsResponse = [{ id: installedAppId }];

    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(installedAppName)}'&$select=id`) {
        return userInstalledAppsResponse;
      }
      throw 'Invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userId)}/teamwork/installedApps/${installedAppId}/upgrade`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: userId, name: installedAppName } });
    assert(postStub.calledOnce);
  });


  it('correctly handles error when app specified by name is not found', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq '${formatting.encodeQueryParameter(installedAppName)}'&$select=id`) {
        return [];
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { userName: userName, name: installedAppName, verbose: true } } as any),
      new CommandError(`The specified Teams app ${installedAppName} does not exist or is not installed for the user`));
  });

  it('correctly handles error when trying to upgrade permanent app', async () => {
    const error = {
      error: {
        code: 'Forbidden',
        message: 'User operation (Upgrade) is not allowed on a permanent app (appId: effe13ef-00e2-4d55-a207-f2d9d1dd7263)',
        innerError: {
          code: 'AccessDenied',
          message: 'User operation (Upgrade) is not allowed on a permanent app (appId: effe13ef-00e2-4d55-a207-f2d9d1dd7263)',
          details: [],
          date: '2024-04-01T18:56:18',
          'request-id': 'edb30b06-1995-46d5-826b-2c9e2e376990',
          'client-request-id': 'edb30b06-1995-46d5-826b-2c9e2e376990'
        }
      }
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${formatting.encodeQueryParameter(userName)}/teamwork/installedApps/${installedAppId}/upgrade`) {
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { userName: userName, id: installedAppId, verbose: true } } as any),
      new CommandError(error.error.message));
  });
});
