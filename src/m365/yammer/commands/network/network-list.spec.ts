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
const command: Command = require('./network-list');

describe(commands.NETWORK_LIST, () => {
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.NETWORK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'email', 'community', 'permalink', 'web_url']);
  });

  it('calls the networking endpoint without parameter', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
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
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 123);
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

  it('calls the networking endpoint without parameter and json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
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
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { debug: true, output: "json" } } as any);
    assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 123);
  });

  it('calls the networking endpoint with parameter', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
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
          ]
        );
      }
      return Promise.reject('Invalid request');
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
