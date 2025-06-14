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
import command from './engage-network-list.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.ENGAGE_NETWORK_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertAccessTokenType').returns();
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
    assert.strictEqual(command.name, commands.ENGAGE_NETWORK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'email', 'community', 'permalink', 'web_url']);
  });

  it('calls the networking endpoint without parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return [
          {
            "id": 123,
            "name": "Network1",
            "email": "email@mail.com",
            "community": true,
            "permalink": "network1-link",
            "web_url": "https://www.yammer.com/network1-link"
          },
          {
            "id": 456,
            "name": "Network2",
            "email": "email2@mail.com",
            "community": false,
            "permalink": "network2-link",
            "web_url": "https://www.yammer.com/network2-link"
          }
        ];
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 123);
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      throw {
        "error": {
          "base": "An error has occurred."
        }
      };
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred.'));
  });

  it('calls the networking endpoint without parameter and json', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return [
          {
            "id": 123,
            "name": "Network1",
            "email": "email@mail.com",
            "community": true,
            "permalink": "network1-link",
            "web_url": "https://www.yammer.com/network1-link"
          },
          {
            "id": 456,
            "name": "Network2",
            "email": "email2@mail.com",
            "community": false,
            "permalink": "network2-link",
            "web_url": "https://www.yammer.com/network2-link"
          }
        ];
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, output: "json" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 123);
  });

  it('calls the networking endpoint with parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return [
          {
            "id": 123,
            "name": "Network1",
            "email": "email@mail.com",
            "community": true,
            "permalink": "network1-link",
            "web_url": "https://www.yammer.com/network1-link"
          },
          {
            "id": 456,
            "name": "Network2",
            "email": "email2@mail.com",
            "community": false,
            "permalink": "network2-link",
            "web_url": "https://www.yammer.com/network2-link"
          }
        ];
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, includeSuspended: true } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 123);
  });

  it('passes validation without parameters', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { includeSuspended: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
