import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './list-get.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.LIST_GET, () => {
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
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
    assert.strictEqual(command.name, commands.LIST_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
      options: commandOptionsSchema.parse({
        name: validName
      })
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
      options: commandOptionsSchema.parse({
        id: validId
      })
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
      options: commandOptionsSchema.parse({
        name: validName
      })
    });

    assert(loggerLogSpy.calledWith(listResponse.value[0]));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw { error: { message: 'An error has occurred' } };
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ id: validId }) } as any), new CommandError('An error has occurred'));
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

});
