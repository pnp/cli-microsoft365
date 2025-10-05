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
import command, { options } from './retentioneventtype-list.js';

describe(commands.RETENTIONEVENTTYPE_LIST, () => {

  //#region Mocked responses
  const mockResponseArray = [
    {
      "displayName": "Retention Event Type",
      "description": "",
      "createdDateTime": "2023-02-02T15:47:54Z",
      "lastModifiedDateTime": "2023-02-02T15:47:54Z",
      "id": "81fa91bd-66cd-4c6c-b0cb-71f37210dc74",
      "createdBy": {
        "user": {
          "id": "36155f4e-bdbd-4101-ba20-5e78f5fba9a9",
          "displayName": null
        }
      },
      "lastModifiedBy": {
        "user": {
          "id": "36155f4e-bdbd-4101-ba20-5e78f5fba9a9",
          "displayName": null
        }
      }
    }
  ];

  const mockResponse = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#security/triggerTypes/retentionEventTypes",
    "@odata.count": 1,
    "value": mockResponseArray
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
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENTTYPE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'createdDateTime']);
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ option: "value" });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves retention event types', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes`) {
        return mockResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });
    assert(loggerLogSpy.calledWith(mockResponseArray));
  });

  it('handles error when retrieving event types', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('An error has occurred'));
  });
});
