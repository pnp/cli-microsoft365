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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./user-get');

describe(commands.USER_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.USER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email', 'job_title', 'state', 'url']);
  });

  it('calls user by e-mail', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/by_email.json?email=pl%40nubo.eu') {
        return Promise.resolve(
          [{ "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { email: "pl@nubo.eu" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
  });

  it('calls user by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/1496550646.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { id: 1496550646 } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
  });

  it('calls the current user and json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/current.json') {
        return Promise.resolve(
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { output: 'json' } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].id, 1496550646);
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

  it('correctly handles 404 error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject({
        "statusCode": 404
      });
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
