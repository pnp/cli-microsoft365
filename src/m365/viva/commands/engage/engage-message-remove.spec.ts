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
import command, { options } from './engage-message-remove.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_MESSAGE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    assert.strictEqual(command.name, commands.ENGAGE_MESSAGE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if id is not provided', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('calls the messaging endpoint with the right parameters and confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 10123190123123, force: true }) });
    assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
  });

  it('calls the messaging endpoint with the right parameters without confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 10123190123123, force: false }) });
    assert.strictEqual(requestDeleteStub.lastCall.args[0].url, 'https://www.yammer.com/api/v1/messages/10123190123123.json');
  });

  it('does not call the messaging endpoint without confirmation', async () => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return;
      }
      throw 'Invalid request';
    });

    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, id: 10123190123123, force: false }) });
    assert(requestDeleteStub.notCalled);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'delete').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: 10123190123123, force: true }) }), new CommandError('An error has occurred.'));
  });
});
