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
import command, { options } from './meeting-transcript-get.js';
import { PassThrough } from 'stream';

describe(commands.MEETING_TRANSCRIPT_GET, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const email = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const id = 'MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh';
  const outputFile = 'transcript.vtt';
  const meetingTranscriptResponse = {
    "id": "MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh",
    "meetingId": "MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy",
    "meetingOrganizerId": "68be84bf-a585-4776-80b3-30aa5207aa21",
    "transcriptContentUrl": "https://graph.microsoft.com/beta/users/68be84bf-a585-4776-80b3-30aa5207aa21/onlineMeetings/MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy/transcripts/MSMjMCMjZDAwYWU3NjUtNmM2Yi00NjQxLTgwMWQtMTkzMmFmMjEzNzdh/content",
    "createdDateTime": "2021-09-17T06:09:24.8968037Z"
  };

  const validOptions = { meetingId: meetingId, id: id };

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
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
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
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get,
      entraUser.getUserIdByEmail,
      entraUser.getUserIdByUpn,
      cli.executeCommandWithOutput,
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

  it('preserves option aliases, types and requirements in command metadata', () => {
    const actual = commandInfo.options
      .filter(option => !['debug', 'verbose', 'output', 'query'].includes(option.name))
      .map(option => ({ name: option.name, short: option.short, type: option.type, required: option.required }));
    assert.deepStrictEqual(actual, [
      { name: 'userId', short: 'u', type: 'string', required: false },
      { name: 'userName', short: 'n', type: 'string', required: false },
      { name: 'email', short: undefined, type: 'string', required: false },
      { name: 'meetingId', short: 'm', type: 'string', required: true },
      { name: 'id', short: 'i', type: 'string', required: true },
      { name: 'outputFile', short: 'f', type: 'string', required: false }
    ]);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ ...validOptions, unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation without a user selector', () => {
    const actual = commandOptionsSchema.safeParse(validOptions);
    assert.strictEqual(actual.success, true);
  });

  for (const [selector, value, token] of [
    ['userId', userId, '@meid'],
    ['userName', userName, '@meusername'],
    ['email', userName, '@meusername']
  ]) {
    it(`passes validation with a valid ${selector}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [selector]: value });
      assert.strictEqual(actual.success, true);
    });

    it(`fails validation with an invalid ${selector}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [selector]: 'invalid' });
      assert.strictEqual(actual.success, false);
    });

    it(`passes validation with the runtime token for ${selector}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [selector]: token });
      assert.strictEqual(actual.success, true);
    });
  }

  for (const selectors of [
    { userId: userId, userName: userName },
    { userId: userId, email: userName },
    { userName: userName, email: userName },
    { userId: userId, userName: userName, email: userName }
  ]) {
    it(`fails validation with conflicting selectors: ${Object.keys(selectors).join(', ')}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, ...selectors });
      assert.strictEqual(actual.success, false);
      assert(actual.error?.issues.some(issue => issue.code === 'custom' &&
        issue.params?.customCode === 'optionSet' &&
        JSON.stringify(issue.params.options) === JSON.stringify(['userId', 'userName', 'email'])));
    });
  }

  for (const requiredOption of Object.keys(validOptions)) {
    it(`fails validation without required option ${requiredOption}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [requiredOption]: undefined });
      assert.strictEqual(actual.success, false);
    });
  }

  it('fails validation when the userId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({ userId: 'foo', meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the userName is not valid', () => {
    const actual = commandOptionsSchema.safeParse({ userName: 'foo', meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the email is not valid', () => {
    const actual = commandOptionsSchema.safeParse({ email: 'foo', meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('succeeds validation when the userId, meetingId, and id are valid', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, true);
  });

  it('succeeds validation when the userName, meetingId, and id are valid', () => {
    const actual = commandOptionsSchema.safeParse({ userName: userName, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, true);
  });

  it('succeeds validation when the email, meetingId, and id are valid', () => {
    const actual = commandOptionsSchema.safeParse({ email: email, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when the userId, email, and userName are given', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId, userName: userName, email: email, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the userId and email are given', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId, email: email, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the userId and userName are given', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId, userName: userName, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the userName and email are given', () => {
    const actual = commandOptionsSchema.safeParse({ userName: userName, email: email, meetingId: meetingId, id: id });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: id, outputFile: 'abc' });
    sinonUtil.restore(fs.existsSync);
    assert.strictEqual(actual.success, false);
  });

  it('retrieves transcript correctly for the given meetingId for the current user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/onlineMeetings/${meetingId}/transcripts/${id}`) {
        return meetingTranscriptResponse;
      }
      throw 'Invalid request.';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: id }) });
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

    await command.action(logger, { options: commandOptionsSchema.parse({ userId: userId, meetingId: meetingId, id: id }) });

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

    await command.action(logger, { options: commandOptionsSchema.parse({ userName: userName, meetingId: meetingId, id: id }) });

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

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, email: email, meetingId: meetingId, id: id }) });
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
      await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, meetingId: meetingId, id: id, outputFile: outputFile }) });
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: id }) }),
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

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: id, outputFile: outputFile }) }),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when throwing request', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, meetingId: meetingId, id: id }) }),
      new CommandError(errorMessage));
  });

  it('correctly handles error when options are missing', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: id }) }),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting transcript using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ userId: userId, meetingId: meetingId, id: id }) }),
      new CommandError(`The options 'userId', 'userName', and 'email' cannot be used while retrieving meeting transcript using delegated permissions`));
  });
});