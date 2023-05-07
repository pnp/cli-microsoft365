import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { formatting } from '../../../../utils/formatting';

const command: Command = require('./meeting-transcript-list');

describe(commands.MEETING_TRANSCRIPT_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const email = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const meetingTranscriptsResponse =
    [
      {
        "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
        "createdDateTime": "2021-09-17T06:09:24.8968037Z"
      },
      {
        "id": "MSMjMCMjMzAxNjNhYTctNWRmZi00MjM3LTg5MGQtNWJhYWZjZTZhNWYw",
        "createdDateTime": "2021-09-16T18:58:58.6760692Z"
      },
      {
        "id": "MSMjMCMjNzU3ODc2ZDYtOTcwMi00MDhkLWFkNDItOTE2ZDNmZjkwZGY4",
        "createdDateTime": "2021-09-16T18:56:00.9038309Z"
      }
    ];

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
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_TRANSCRIPT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'createdDateTime']);
  });

  it('fails validation when the userId is not a GUID', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userName is not valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userName: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('succeeds validation when the userId and meetingId are valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('succeeds validation when the userName and meetingId are valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('succeeds validation when the email and meetingId are valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, email: email } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the userId and email and userName are given', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, userName: userName, email: email } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when given email is not valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, email: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and email are given', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, email: email } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and userName are given', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId, userName: userName } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('retrieves transcript list correctly for the given meetingId for the current user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts`) {
        return { value: meetingTranscriptsResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      {
        meetingId: meetingId
      }
    });

    assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
  });

  it('retrieves meeting transcript list correctly by meetingId and userID', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts`) {
        return { value: meetingTranscriptsResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      {
        meetingId: meetingId,
        userId: userId
      }
    });

    assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
  });

  it('retrieves meeting transcript list correctly by meetingId and userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/onlineMeetings/${meetingId}/transcripts`) {
        return { value: meetingTranscriptsResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      {
        meetingId: meetingId,
        userName: userName
      }
    });

    assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
  });

  it('retrieves meeting transcript list correctly by meetingId and email', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(email)}'&$select=id`) {
        return {
          value: [
            {
              id: userId
            }]
        };
      }

      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts`) {
        return { value: meetingTranscriptsResponse };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      {
        meetingId: meetingId,
        email: email,
        verbose: true
      }
    });

    assert(loggerLogSpy.calledWith(meetingTranscriptsResponse));
  });

  it('correctly handles error when throwing request', async () => {
    const errorMessage = 'An error has occured';

    sinon.stub(request, 'get').callsFake(async () => {
      throw { error: { error: { message: errorMessage } } };
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, meetingId: meetingId } } as any),
      new CommandError(errorMessage));
  });

  it('correctly handles error when options are missing', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => true);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting transcripts list using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').callsFake(() => false);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, userId: userId } } as any),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used while retrieving meeting transcripts using delegated permissions`));
  });
});