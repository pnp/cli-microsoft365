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

import command from './message-like-set.js';
describe(commands.MESSAGE_LIKE_SET, () => {
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let requests: any[];
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
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
    requests = [];
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MESSAGE_LIKE_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').rejects({
      "error": {
        "base": "An error has occurred."
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { messageId: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { messageId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if enabled set to "true"', async () => {
    const actual = await command.validate({ options: { messageId: 10123123, enable: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if enabled set to "false"', async () => {
    const actual = await command.validate({ options: { messageId: 10123123, enable: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts when confirmation argument not passed', async () => {
    await command.action(logger, { options: { messageId: 1231231, enable: false } });


    assert(promptIssued);
  });

  it('calls the service when liking a message', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, messageId: 1231231 } });
    assert(requestPostedStub.called);
  });

  it('calls the service when liking a message and confirm passed', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, messageId: 1231231, force: true } });
    assert(requestPostedStub.called);
  });

  it('calls the service when liking a message and enabled set to true', async () => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, messageId: 1231231, enable: true } });
    assert(requestPostedStub.called);
  });

  it('calls the service when disliking a message and confirming', async () => {
    const requestPostedStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, messageId: 1231231, enable: false, force: true } });
    assert(requestPostedStub.called);
  });

  it('prompts when disliking and confirmation parameter is denied', async () => {
    await command.action(logger, { options: { messageId: 1231231, enable: false, force: false } });


    assert(promptIssued);
  });

  it('calls the service when disliking a message and confirmation is hit', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { debug: true, messageId: 1231231, enable: false } });
    assert(requestDeleteStub.called);
  });

  it('Aborts execution when enabled set to false and confirmation is not given', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { messageId: 1231231, enable: false } });
    assert(requests.length === 0);
  });
}); 