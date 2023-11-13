import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { formatting } from '../../../../utils/formatting.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './user-app-remove.js';

describe(commands.USER_APP_REMOVE, () => {
  const userId = '15d7a78e-fd77-4599-97a5-dbb6372846c6';
  const userName = 'contoso@contoso.onmicrosoft.com';
  const appId = 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=';
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;

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
    (command as any).items = [];
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.USER_APP_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        userId: 'invalid',
        id: appId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userName is not a valid UPN.', async () => {
    const actual = await command.validate({
      options: {
        userName: "no-an-email",
        id: appId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', async () => {
    const actual = await command.validate({
      options: {
        id: appId,
        userId: userId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the input is correct (userName)', async () => {
    const actual = await command.validate({
      options: {
        id: appId,
        userName: userName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the app when confirmation argument is not passed', async () => {
    await command.action(logger, {
      options: {
        userId: userId,
        id: appId
      }
    } as any);

    assert(promptIssued);
  });

  it('aborts removing the app by id when confirmation prompt is not continued', async () => {
    const requestDeleteSpy = sinon.stub(request, 'delete');

    await command.action(logger, {
      options: {
        userId: userId,
        id: appId
      }
    } as any);
    assert(requestDeleteSpy.notCalled);
  });

  it('removes the app by id for the specified user when confirmation is specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/users/${userId}/teamwork/installedApps/${appId}`) > -1) {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userId: userId,
        id: appId,
        debug: true,
        force: true
      }
    } as any);
  });

  it('removes the app by id for the specified user using username when confirmation is specified.', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/users/${formatting.encodeQueryParameter(userName)}/teamwork/installedApps/${appId}`) > -1) {
        return Promise.resolve();
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userName: userName,
        id: appId,
        force: true
      }
    } as any);
  });

  it('removes the app by id for the specified user when prompt is confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/users/${userId}/teamwork/installedApps/${appId}`) > -1) {
        return Promise.resolve();
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        userId: userId,
        id: appId,
        debug: true
      }
    } as any);
  });

  it('removes the app for the specified user using username', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/users/${userId}/teamwork/installedApps/${appId}`) > -1) {
        return Promise.resolve();
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        userName: userName,
        id: appId
      }
    } as any);
  });

  it('removes the app by name for the specified user when prompt is confirmed (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq 'TeamsApp'`) {
        return {
          "value": [
            {
              "id": "YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=",
              "displayName": "TeamsApp"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) {
        return Promise.resolve();
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'TeamsApp',
        debug: true
      }
    } as any);
  });

  it('fails to get teams app when app by name does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq 'TeamsApp'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        name: 'TeamsApp',
        force: true
      }
    } as any), new CommandError('The specified Teams app does not exist'));
  });

  it('handles selecting single result when multiple teams apps with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps?$expand=teamsAppDefinition&$filter=teamsAppDefinition/displayName eq 'TeamsApp'`) {
        return {
          "value": [
            {
              "id": "ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ=",
              "displayName": "TeamsApp"
            },
            {
              "id": "NmY0ODM2N2EtMjVmMC00NjNmLTlmMGQtMmFiZTBiYmYzNzRjIyMxLjAuMCMjUHVibGlzaGVk",
              "displayName": "TeamsApp"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({ id: "ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ=" });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/ZDczZWVjZmQtYzFkNS00MzY2LWJkMjEtZDUyOTM1ZThkYjkxIyMxLjYuMC4wIyNQdWJsaXNoZWQ=`) {
        removeRequestIssued = true;
        return Promise.resolve();
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: 'c527a470-a882-481c-981c-ee6efaba85c7', name: 'TeamsApp', force: true } });
    assert(removeRequestIssued);
  });

  it('correctly handles error while removing teams app', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/users/${userId}/teamwork/installedApps/${appId}`) > -1) {
        throw error;
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        userId: userId,
        id: appId,
        force: true
      }
    } as any), new CommandError(error.error.message));
  });
});
