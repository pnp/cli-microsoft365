import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./retentionevent-add');

describe(commands.RETENTIONEVENT_ADD, () => {
  const validDisplayName = "Event display name";
  const validDescription = "Event description";
  const validAssetIds = "filesQuery,filesQuery1";
  const validKeyswords = "messagesQuery,messagesQuery1";
  const validDate = "2023-04-02T15:47:54Z";
  const validTypeId = "81fa91bd-66cd-4c6c-b0cb-71f37210dc74";
  const validTypeName = "Event type";
  const EventResponse = {
    "displayName": "Event display name",
    "description": "Event description",
    "eventTriggerDateTime": "2023-04-02T15:47:54Z",
    "lastStatusUpdateDateTime": "0001-01-01T00:00:00Z",
    "createdDateTime": "2023-02-20T18:53:05Z",
    "lastModifiedDateTime": "2023-02-20T18:53:05Z",
    "id": "9f5c1a04-8f7a-4bff-e400-08db1373b324",
    "eventQueries": [
      {
        "queryType": "files",
        "query": "filesQuery,filesQuery1"
      },
      {
        "queryType": "messages",
        "query": "messagesQuery,messagesQuery1"
      }
    ],
    "eventStatus": {
      "error": null,
      "status": "pending"
    },
    "eventPropagationResults": [],
    "createdBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    },
    "lastModifiedBy": {
      "user": {
        "id": null,
        "displayName": "John Doe"
      }
    }
  };

  const eventTypeResponse = {
    value: [
      {
        "displayName": validTypeName,
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
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
      request.post,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.RETENTIONEVENT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if date is not a valid ISO date string', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName, eventTypeId: validTypeId, description: validDescription, triggerDateTime: "Not a valid date", assetIds: validAssetIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if assetId or keywords is not provided', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName, eventTypeId: validTypeId, description: validDescription, triggerDateTime: validDate } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct ISO date string is entered', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName, eventTypeId: validTypeId, description: validDescription, triggerDateTime: validDate, assetIds: validAssetIds } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds retention event with minimal required parameters and assetIds', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: validDisplayName, eventTypeId: validTypeId, assetIds: validAssetIds } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('adds retention event with minimal required parameters and keywords', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: validDisplayName, eventTypeId: validTypeId, keywords: validKeyswords } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('adds retention event with all parameters', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, displayName: validDisplayName, eventTypeId: validTypeId, description: validDescription, triggerDateTime: validDate, assetIds: validAssetIds, keywords: validKeyswords } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('adds retention event with minimal required parameters and assetIds based on event type name', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes`) {
        return eventTypeResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, displayName: validDisplayName, eventTypeName: validTypeName, assetIds: validAssetIds } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('throws error when no event type found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/security/triggerTypes/retentionEventTypes`) {
        return ({ "value": [] });
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: { displayName: validDisplayName, eventTypeName: validTypeName, assetIds: validAssetIds }
    }), new CommandError(`The specified event type '${validTypeName}' does not exist.`));
  });

  it('correctly handles random API error', async () => {
    const error = {
      error: {
        message: 'The purview retention event cannot be added.'
      }
    };
    sinon.stub(request, 'post').callsFake(async () => { throw error; });

    await assert.rejects(command.action(logger, {
      options: {
        displayName: validDisplayName, eventTypeId: validTypeId, assetIds: validAssetIds
      }
    }), new CommandError(error.error.message));
  });
});