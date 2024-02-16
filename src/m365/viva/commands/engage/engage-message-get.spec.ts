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
import command from './engage-message-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import yammerCommands from './yammerCommands.js';


describe(commands.ENGAGE_MESSAGE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const firstMessage: any = { "sender_id": 1496550646, "replied_to_id": 1496550647, "id": 10123190123123, "thread_id": "", group_id: 11231123123, created_at: "2019/09/09 07:53:18 +0000", "content_excerpt": "message1" };
  const secondMessage: any = { "sender_id": 1496550640, "replied_to_id": "", "id": 10123190123124, "thread_id": "", group_id: "", created_at: "2019/09/08 07:53:18 +0000", "content_excerpt": "message2" };

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_MESSAGE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [yammerCommands.MESSAGE_GET]);
  });

  it('correctly logs deprecation warning for yammer command', async () => {
    const chalk = (await import('chalk')).default;
    const loggerErrSpy = sinon.spy(logger, 'logToStderr');
    const commandNameStub = sinon.stub(cli, 'currentCommandName').value(yammerCommands.MESSAGE_GET);
    sinon.stub(request, 'get').resolves(firstMessage);

    await command.action(logger, { options: { id: 10123190123123 } });
    assert.deepStrictEqual(loggerErrSpy.firstCall.firstArg, chalk.yellow(`Command '${yammerCommands.MESSAGE_GET}' is deprecated. Please use '${commands.ENGAGE_MESSAGE_GET}' instead.`));

    sinonUtil.restore([loggerErrSpy, commandNameStub]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'sender_id', 'replied_to_id', 'thread_id', 'group_id', 'created_at', 'direct_message', 'system_message', 'privacy', 'message_type', 'content_excerpt']);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { id: 'nonumber' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('id is required', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('calls the messaging endpoint with the right parameters', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return firstMessage;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: 10123190123123, debug: true } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 10123190123123);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('calls the messaging endpoint with id and json and json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123124.json') {
        return secondMessage;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true, id: 10123190123124, output: "json" } } as any);

    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 10123190123124);
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { id: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
