import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './retentioneventtype-add.js';

describe(commands.RETENTIONEVENTTYPE_ADD, () => {
  const displayName = 'Contract Expiry';
  const description = 'A retention event type description';

  //#region Mocked Responses
  const requestResponse = {
    displayName: displayName,
    description: description,
    createdDateTime: "2022-12-21T09:28:37Z",
    lastModifiedDateTime: "2022-12-21T09:28:37Z",
    id: "f7e05955-210b-4a8e-a5de-3c64cfa6d9be",
    createdBy: {
      user: {
        id: null,
        displayName: "John Doe"
      }
    },
    lastModifiedBy: {
      user: {
        id: null,
        displayName: "John Doe"
      }
    }
  };
  //#endregion

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
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds retention event type', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes`) {
        return requestResponse;
      }

      return 'Invalid Request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: displayName }) });
    assert(loggerLogSpy.calledWith(requestResponse));
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ displayName: displayName, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      throw 'An error has occurred.';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: displayName }) }),
      new CommandError('An error has occurred.'));
  });
});