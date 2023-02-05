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
const command: Command = require('./retentionevent-remove');

describe(commands.RETENTIONEVENT_REMOVE, () => {
  const validId = 'c37d695e-d581-4ae9-82a0-9364eba4291e';

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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
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
    assert.strictEqual(command.name, commands.RETENTIONEVENT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified retention event when confirm option not passed', async () => {
    await command.action(logger, {
      options: {
        id: validId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified retention event when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        id: validId
      }
    });
    assert(deleteSpy.notCalled);
  });

  it('Correctly deletes retention event by id', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents/${validId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        id: validId
      }
    });
  });

  it('Correctly deletes retention event by id when prompt confirmed', async () => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents/${validId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        id: validId,
        confirm: true
      }
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'delete').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        id: validId,
        confirm: true
      }
    }), new CommandError("An error has occurred"));
  });
});