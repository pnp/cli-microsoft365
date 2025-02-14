import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './meeting-transcript-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import { PassThrough } from 'stream';

describe(commands.MEETING_TRANSCRIPT_GET, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const email = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const id = 'MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh';
  const outputFile = 'c:\transcript.vtt';
  const meetingTranscriptResponse = {
    "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
    "meetingId": "MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy",
    "meetingOrganizerId": "68be84bf-a585-4776-80b3-30aa5207aa21",
    "transcriptContentUrl": "https://graph.microsoft.com/beta/users/68be84bf-a585-4776-80b3-30aa5207aa21/onlineMeetings/MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy/transcripts/MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh/content",
    "createdDateTime": "2021-09-17T06:09:24.8968037Z"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get,
      entraUser.getUserIdByEmail,
      entraUser.getUserIdByUpn,
      cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue,
      fs.createWriteStream
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_TRANSCRIPT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when the userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'foo', meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userName is not valid', async () => {
    const actual = await command.validate({ options: { userName: 'foo', meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the email is not valid', async () => {
    const actual = await command.validate({ options: { email: 'foo', meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('succeeds validation when the userId, meetingId, and id are valid', async () => {
    const actual = await command.validate({ options: { userId: userId, meetingId: meetingId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('succeeds validation when the userName, meetingId, and id are valid', async () => {
    const actual = await command.validate({ options: { userName: userName, meetingId: meetingId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('succeeds validation when the email, meetingId, and id are valid', async () => {
    const actual = await command.validate({ options: { email: email, meetingId: meetingId, id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the userId, email, and userName are given', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId, userName: userName, email: email, meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and email are given', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId, email: email, meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId and userName are given', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId, userName: userName, meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userName and email are given', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { userId: userId, email: email, meetingId: meetingId, id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if path doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({ options: { meetingId: meetingId, id: id, outputFile: 'abc' } }, commandInfo);
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('retrieves transcript correctly for the given meetingId for the current user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return meetingTranscriptResponse;
      }
      throw 'Invalid request.';
    });

    await command.action(logger, { options: { meetingId: meetingId, id: id } });
    assert(loggerLogSpy.calledWith(meetingTranscriptResponse));
  });

  it('retrieves transcript correctly for the given id, meetingId, and userID', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return meetingTranscriptResponse;
      }

      throw 'Invalid request.';
    });

    await command.action(logger, { options: { userId: userId, meetingId: meetingId, id: id } });

    assert(loggerLogSpy.calledWith(meetingTranscriptResponse));
  });

  it('retrieves transcript correctly for the given id, meetingId, and userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return meetingTranscriptResponse;
      }

      throw 'Invalid request.';
    });

    await command.action(logger, { options: { userName: userName, meetingId: meetingId, id: id } });

    assert(loggerLogSpy.calledWith(meetingTranscriptResponse));
  });

  it('retrieves transcript correctly for the given id, meetingId, and email (verbose)', async () => {
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

      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return meetingTranscriptResponse;
      }

      throw 'Invalid request.';
    });

    await command.action(logger, { options: { verbose: true, email: email, meetingId: meetingId, id: id } });
    assert(loggerLogSpy.calledWith(meetingTranscriptResponse));
  });

  it('downloads a transcript when outputFile is specified (verbose)', async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 0);

    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts/${id}/content?$format=text/vtt`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request.';
    });

    try {
      await command.action(logger, { options: { verbose: true, meetingId: meetingId, id: id, outputFile: outputFile } });
      assert(fsStub.calledOnce);
    }
    finally {
      sinonUtil.restore([
        fs.createWriteStream
      ]);
    }
  });

  it('correctly handles error when the meeting transcript not found', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return;
      }

      throw 'The specified meeting transcript was not found';
    });

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: id } }),
      new CommandError(`The specified meeting transcript was not found`));
  });

  it(`handles error when saving the transcript to file fails`, async () => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('error', "An error has occurred");
    }, 0);

    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts/${id}/content?$format=text/vtt`) {
        return {
          data: responseStream
        };
      }

      throw 'Invalid request.';
    });

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: id, outputFile: outputFile } }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when throwing request', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: { verbose: true, meetingId: meetingId, id: id } } as any),
      new CommandError(errorMessage));
  });

  it('correctly handles error when options are missing', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: id } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting transcript using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: { userId: userId, meetingId: meetingId, id: id } } as any),
      new CommandError(`The options 'userId', 'userName', and 'email' cannot be used while retrieving meeting transcript using delegated permissions`));
  });
});