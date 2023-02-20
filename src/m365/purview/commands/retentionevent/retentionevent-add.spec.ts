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
import { accessToken } from '../../../../utils/accessToken';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./retentionevent-add');

describe(commands.RETENTIONEVENT_ADD, () => {
  const validDisplayName = "Event display name";
  const validDescription = "Event description";
  const validAssetIds = "filesQuery,filesQuery1";
  const validKeyswords = "messagesQuery,messagesQuery1";
  const validDate = "2023-04-02T15:47:54Z";
  const validType = "Event type";
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
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
    const actual = await command.validate({ options: { displayName: validDisplayName, eventType: validType, description: validDescription, triggerDateTime: "Not a valid date", assetIds: validAssetIds } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if assetId or keywords is not provided', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName, eventType: validType, description: validDescription, triggerDateTime: validDate } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if a correct ISO date string is entered', async () => {
    const actual = await command.validate({ options: { displayName: validDisplayName, eventType: validType, description: validDescription, triggerDateTime: validDate, assetIds: validAssetIds } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds retention event with minimal required parameters and assetIds', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: validDisplayName, eventType: validType, assetIds: validAssetIds } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('adds retention event with minimal required parameters and keywords', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: validDisplayName, eventType: validType, keywords: validKeyswords } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('adds retention event with all parameters', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/security/triggers/retentionEvents`) {
        return EventResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, displayName: validDisplayName, eventType: validType, description: validDescription, triggerDateTime: validDate, assetIds: validAssetIds, keywords: validKeyswords } });
    assert(loggerLogSpy.calledWith(EventResponse));
  });

  it('throws error if something fails using application permissions', async () => {
    sinonUtil.restore([accessToken.isAppOnlyAccessToken]);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`This command does not support application permissions.`));
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
        displayName: validDisplayName, eventType: validType, assetIds: validAssetIds
      }
    }), new CommandError(error.error.message));
  });
});