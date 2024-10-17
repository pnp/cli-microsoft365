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
import command from './engage-user-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_USER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
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

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email', 'job_title', 'state', 'url']);
  });

  it('calls user by e-mail', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/by_email.json?email=pl%40nubo.eu') {
        return [{ "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" }];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { email: "pl@nubo.eu" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('calls user by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/1496550646.json') {
        return { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { id: 1496550646 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
  });

  it('calls the current user and json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/current.json') {
        return { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "John Doe" };
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw { "error": { "base": "An error has occurred." } };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('correctly handles 404 error', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw {
        "statusCode": 404
      };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Not found (404)'));
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if id set ', async () => {
    const actual = await command.validate({ options: { id: 1496550646 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if email set', async () => {
    const actual = await command.validate({ options: { email: "pl@nubo.eu" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not pass with id and e-mail', async () => {
    const actual = await command.validate({ options: { id: 1496550646, email: "pl@nubo.eu" } }, commandInfo);
    assert.strictEqual(actual, "You are only allowed to search by ID or e-mail but not both");
  });
});
