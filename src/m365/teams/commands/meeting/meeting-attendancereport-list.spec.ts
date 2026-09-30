import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import userGetCommand from '../../../entra/commands/user/user-get.js';
import commands from '../../commands.js';
import command, { options } from './meeting-attendancereport-list.js';

describe(commands.MEETING_ATTENDANCEREPORT_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'user@tenant.com';
  const meetingId = 'MSo5MWZmMmUxNy04NGRlLTQ1NWEtODgxNS01MmIyMTY4M2Y2NGUqMCoqMTk6bWVldGluZ19ZMlEzTlRRMFpEWXRaamMzWkMwMFlUVmhMVGt4TTJJdFpURmtNMkUwTUdGak1qVmpAdGhyZWFkLnYy';
  const meetingAttendanceResponse =
    [
      {
        "id": "ae6ddf54-5d48-4448-a7a9-780eee17fa13",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:46:46.981Z",
        "meetingEndDateTime": "2022-11-22T22:47:07.703Z"
      },
      {
        "id": "3fd019cc-6df5-485f-86a0-96838ab98e66",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:45:10.226Z",
        "meetingEndDateTime": "2022-11-22T22:45:22.347Z"
      },
      {
        "id": "04ddf3a5-0c02-4865-928e-9b65d1b33570",
        "totalParticipantCount": 1,
        "meetingStartDateTime": "2022-11-22T22:43:38.052Z",
        "meetingEndDateTime": "2022-11-22T22:44:12.893Z"
      }
    ];

  const validOptions = { meetingId: meetingId };

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
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_ATTENDANCEREPORT_LIST);
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
      { name: 'meetingId', short: 'm', type: 'string', required: true }
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

    it(`passes validation with the runtime token for ${selector}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [selector]: token });
      assert.strictEqual(actual.success, true);
    });
  }

  for (const selector of ['userName', 'email']) {
    it(`preserves acceptance of unvalidated ${selector} values`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [selector]: 'invalid' });
      assert.strictEqual(actual.success, true);
    });
  }

  for (const selectors of [
    { userId: userId, userName: userName },
    { userId: userId, email: userName },
    { userName: userName, email: userName },
    { userId: userId, userName: userName, email: userName }
  ]) {
    it(`preserves acceptance of multiple selectors: ${Object.keys(selectors).join(', ')}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, ...selectors });
      assert.strictEqual(actual.success, true);
    });
  }

  for (const requiredOption of Object.keys(validOptions)) {
    it(`fails validation without required option ${requiredOption}`, () => {
      const actual = commandOptionsSchema.safeParse({ ...validOptions, [requiredOption]: undefined });
      assert.strictEqual(actual.success, false);
    });
  }

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'totalParticipantCount']);
  });

  it('fails validation when the userId is not a GUID', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, userId: 'foo' });
    assert.strictEqual(actual.success, false);
  });

  it('succeeds validation when the userId and meetingId are valid', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves meeting attendace reports correctly for the current user', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        meetingId: meetingId
      })
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userId', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        meetingId: meetingId,
        userId: userId
      })
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
      }
      throw 'Invalid request.';
    });

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { stdout: JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        meetingId: meetingId,
        userName: userName
      })
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userEmail', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { stdout: JSON.stringify({ id: userId }) };
      }
      throw 'Invalid request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
      }
      throw 'Invalid request.';
    });

    await command.action(logger, {
      options:
      commandOptionsSchema.parse({
        meetingId: meetingId,
        email: userName,
        verbose: true
      })
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('correctly handles error when throwing request', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').rejects({ error: { error: { message: errorMessage } } });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, meetingId: meetingId }) }),
      new CommandError(errorMessage));
  });

  it('correctly handles error when options are missing', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId }) }),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, userId: userId }) }),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`));
  });
});
