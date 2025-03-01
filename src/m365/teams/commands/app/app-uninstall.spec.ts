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
import command from './app-uninstall.js';

describe(commands.APP_UNINSTALL, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.APP_UNINSTALL);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '123456789',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('uninstalls an app from a Microsoft Teams team with confirmation', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        force: true,
        verbose: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('uninstalls an app from a Microsoft Teams team without confirmation', async () => {
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/c527a470-a882-481c-981c-ee6efaba85c7/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('aborts uninstalling an app from a Microsoft Teams team when prompt not confirmed', async () => {
    sinon.stub(cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error when uninstalling an app', async () => {
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
    sinon.stub(request, 'delete').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        force: true
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        force: true
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
