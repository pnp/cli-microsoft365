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
import command, { options } from './engage-user-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_USER_LIST, () => {
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
    assert.strictEqual(command.name, commands.ENGAGE_USER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email']);
  });

  it('passes validation without parameters', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('passes validation with parameters', () => {
    const actual = commandOptionsSchema.safeParse({ letter: "A" });
    assert.strictEqual(actual.success, true);
  });

  it('letter does not allow numbers', () => {
    const actual = commandOptionsSchema.safeParse({ letter: "1" });
    assert.strictEqual(actual.success, false);
  });

  it('groupId must be a number', () => {
    const actual = commandOptionsSchema.safeParse({ groupId: "aasdf" });
    assert.strictEqual(actual.success, false);
  });

  it('limit must be a number', () => {
    const actual = commandOptionsSchema.safeParse({ limit: "aasdf" });
    assert.strictEqual(actual.success, false);
  });

  it('sortBy validation check', () => {
    const actual = commandOptionsSchema.safeParse({ sortBy: "aasdf" });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if letter is set to a single character', () => {
    const actual = commandOptionsSchema.safeParse({ letter: "a" });
    assert.strictEqual(actual.success, true);
  });

  it('does not pass validation if letter is set to a multiple characters', () => {
    const actual = commandOptionsSchema.safeParse({ letter: "ab" });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ id: 1, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('returns all network users', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return [
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('returns all network users using json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return [
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }
        ];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('sorts network users by messages', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&sort_by=messages') {
        return [
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }
        ];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ sortBy: "messages" }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
  });

  it('fakes the return of more results', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(async () => {
      if (i++ === 0) {
        return {
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }],
          more_available: true
        };
      }
      else {
        return {
          users: [
            { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
            { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }],
          more_available: false
        };
      }
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 4);
  });

  it('fakes the return of more than 50 entries', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(async () => {
      if (i++ === 0) {
        return [
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }];
      }
      else {
        return [
          { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
          { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }];
      }
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ output: 'json' }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 52);
  });

  it('fakes the return of more results with exception', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(async () => {
      if (i++ === 0) {
        return {
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }],
          more_available: true
        };
      }
      else {
        throw ({
          "error": {
            "base": "An error has occurred."
          }
        });
      }
    });
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('An error has occurred.'));
  });

  it('sorts users in reverse order', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&reverse=true') {
        return [
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ reverse: true }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
  });

  it('sorts users in reverse order in a group and limits the user to 2', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1&reverse=true') {
        return {
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }],
          has_more: true
        };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: 5785177, reverse: true, limit: 2 }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 2);
  });

  it('returns users of a specific group', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1') {
        return {
          users: [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }, { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }],
          has_more: false
        };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ groupId: 5785177 }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('returns users starting with the letter P', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&letter=P') {
        return [
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: commandOptionsSchema.parse({ letter: "P" }) });
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw ({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('An error has occurred.'));
  });
});
