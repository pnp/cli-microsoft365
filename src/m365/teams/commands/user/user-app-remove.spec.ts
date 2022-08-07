import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-app-remove');

describe(commands.USER_APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
    (command as any).items = [];
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        userId: 'invalid',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', async () => {
    const actual = await command.validate({
      options: {
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the app when confirmation argument is not passed', async () => {
    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    } as any);
    let promptIssued = false;
    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
  });

  it('aborts removing the app when confirmation prompt is not continued', async () => {
    const requestDeleteSpy = sinon.stub(request, 'delete');

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    } as any);
    assert(requestDeleteSpy.notCalled);
  });

  it('removes the app for the specified user when confirmation is specified (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        debug: true,
        confirm: true
      }
    } as any);
  });

  it('removes the app for the specified user when prompt is confirmed (debug)', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));
    await command.action(logger, {
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        debug: true
      }
    } as any);
  });

  it('correctly handles error while removing teams app', async () => {
    const error = 'An error has occurred';
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.reject(error);
      }
      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, { options: {
      userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
      id: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
      confirm: true } } as any), new CommandError(error));
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