import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-list');

describe(commands.USER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email']);
  });

  it('returns all network users', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: {} } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('returns all network users using json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('sorts network users by messages', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&sort_by=messages') {
        return Promise.resolve([
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }
        ]);
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { sortBy: "messages" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
  });

  it('fakes the return of more results', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }],
          more_available: true
        });
      }
      else {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
            { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }],
          more_available: false
        });
      }
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 4);
  });

  it('fakes the return of more than 50 entries', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve(
          [
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
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }]
        );
      }
      else {
        return Promise.resolve([
          { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
          { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }]);
      }
    });
    await command.action(logger, { options: { output: 'debug' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 52);
  });

  it('fakes the return of more results with exception', async () => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }],
          more_available: true
        });
      }
      else {
        return Promise.reject({
          "error": {
            "base": "An error has occurred."
          }
        });
      }
    });
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('sorts users in reverse order', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&reverse=true') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { reverse: true } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
  });

  it('sorts users in reverse order in a group and limits the user to 2', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1&reverse=true') {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }],
          has_more: true
        });
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { groupId: 5785177, reverse: true, limit: 2 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 2);
  });

  it('returns users of a specific group', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1') {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }, { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }],
          has_more: false
        });
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { groupId: 5785177 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('returns users starting with the letter P', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&letter=P') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Adam Doe" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { letter: "P" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { letter: "A" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('letter does not allow numbers', async () => {
    const actual = await command.validate({ options: { letter: "1" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', async () => {
    const actual = await command.validate({ options: { groupId: "aasdf" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('limit must be a number', async () => {
    const actual = await command.validate({ options: { limit: "aasdf" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('sortBy validation check', async () => {
    const actual = await command.validate({ options: { sortBy: "aasdf" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if letter is set to a single character', async () => {
    const actual = await command.validate({ options: { letter: "a" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not pass validation if letter is set to a multiple characters', async () => {
    const actual = await command.validate({ options: { letter: "ab" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
