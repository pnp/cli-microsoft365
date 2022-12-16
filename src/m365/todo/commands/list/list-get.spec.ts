import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-get');

describe(commands.LIST_GET, () => {
  let commandInfo: CommandInfo;
  const validName: string = "Task list";
  const validId: string = "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA=";
  const listResponse = {
    value: [
      {
        displayName: "test cli",
        isOwner: true,
        isShared: false,
        wellknownListName: "none",
        id: "AAMkADY3NmM5ZjhiLTc3M2ItNDg5ZC1iNGRiLTAyM2FmMjVjZmUzOQAuAAAAAAAZ1T9YqZrvS66KkevskFAXAQBEMhhN5VK7RaaKpIc1KhMKAAAZ3e1AAAA="
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    assert.strictEqual(command.name.startsWith(commands.LIST_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'id']);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (name)', async () => {
    const actual = await command.validate({ options: { name: validName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws an error when no list found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(validName)}'`)) {
        if ((opts.headers?.accept as string)?.indexOf('application/json') === 0) {
          return ({ value: [] });
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validName
      }
    }), new CommandError(`The specified list '${validName}' does not exist.`));
  });

  it('lists a specific To Do task list based on the id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/${validId}`) {
        return (listResponse.value[0]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: validId
      }
    });

    assert(loggerLogSpy.calledWith(listResponse.value[0]));
  });

  it('lists a specific To Do task list based on the name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq '${formatting.encodeQueryParameter(validName)}'`) {
        return (listResponse);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: validName
      }
    });

    assert(loggerLogSpy.calledWith(listResponse.value[0]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw { error: { message: 'An error has occurred' } };
    });

    await assert.rejects(command.action(logger, { options: { id: validId } } as any), new CommandError('An error has occurred'));
  });
});
