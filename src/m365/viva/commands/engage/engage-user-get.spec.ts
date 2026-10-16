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
import command, { options } from './engage-user-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_USER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_USER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation without parameters', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if id set', () => {
    const actual = commandOptionsSchema.safeParse({ id: 1496550646 });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if email set', () => {
    const actual = commandOptionsSchema.safeParse({ email: "pl@nubo.eu" });
    assert.strictEqual(actual.success, true);
  });

  it('does not pass with id and e-mail', () => {
    const actual = commandOptionsSchema.safeParse({ id: 1496550646, email: "pl@nubo.eu" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('calls user by e-mail', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/by_email.json?email=pl%40nubo.eu') {
        return [{ "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ email: "pl@nubo.eu" }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('calls user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/1496550646.json') {
        return { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ id: 1496550646 }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
  });

  it('calls the current user and json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/current.json') {
        return { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw { "error": { "base": "An error has occurred." } };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('An error has occurred.'));
  });

  it('correctly handles 404 error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        "statusCode": 404
      };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('Not found (404)'));
  });
});
