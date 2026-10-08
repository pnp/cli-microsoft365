import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './meeting-attendancereport-get.js';
import { entraUser } from '../../../../utils/entraUser.js';
import request from '../../../../request.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MEETING_ATTENDANCEREPORT_GET, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const userName = 'john@contoso.com';
  const meetingId = 'MSpmZTM2Zjc1ZS1jMTAzLTQxMGItYTE4YS0yYmY2ZGYwNmFjM2EqMCoqMTk6bWVldGluZ19NRGt4TnpSaE56UXRZekZtWlMwMFlqWTFMVGhoTVRFdFpUWTBOV1JqTnpoaFkyVTVAdGhyZWFkLnYy';
  const attendanceReportId = 'a8634e64-3147-4a56-9b19-cc822e9c7972';

  const response = {
    id: 'a8634e64-3147-4a56-9b19-cc822e9c7972',
    totalParticipantCount: 1,
    meetingStartDateTime: '2024-04-06T08:18:21.668Z',
    meetingEndDateTime: '2024-04-06T08:18:28.482Z',
    attendanceRecords: [
      {
        id: userId,
        emailAddress: 'john@contoso.com',
        totalAttendanceInSeconds: 3,
        role: 'Organizer',
        identity: {
          id: userId,
          displayName: 'John Doe',
          tenantId: 'e1dd4023-a656-480a-8a0e-b1b1eec51e1e'
        },
        attendanceIntervals: [
          {
            joinDateTime: '2024-04-06T08:18:24.5069531Z',
            leaveDateTime: '2024-04-06T08:18:28.4820462Z',
            durationInSeconds: 3
          }
        ]
      }
    ]
  };

  const error = {
    error: {
      code: 'BadRequest',
      message: 'Index was out of range. Must be non-negative and less than the size of the collection.\r\nParameter name: index',
      innerError: {
        date: '2024-04-09T19:40:32',
        'request-id': '9b23d2c2-bee8-43e1-b854-1eeba77a562d',
        'client-request-id': '9b23d2c2-bee8-43e1-b854-1eeba77a562d'
      }
    }
  };

  const validOptions = { meetingId: meetingId, id: attendanceReportId };

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
    sinon.stub(entraUser, 'getUserIdByEmail').resolves(userId);
    sinon.stub(entraUser, 'getUserIdByUpn').resolves(userId);
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
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

    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_ATTENDANCEREPORT_GET);
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
      { name: 'id', short: 'i', type: 'string', required: true }
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

  it('retrieves attendance report for currently signed in user', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by userId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, userId: userId, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by userName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, userName: userName, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by email', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true }) });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('throws error when using application permissions and not mentioning userId, userName or email', async () => {
    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, verbose: true }) }),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions.`));
  });

  it('throws error when using delegated permissions and mentioning userId, userName or email', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, userId: userId, verbose: true }) }),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance report using delegated permissions.`));
  });

  it('throws error when meeting not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true }) }),
      new CommandError(error.error.message));
  });

  it('throws error when attendanceReport not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true }) }),
      new CommandError(error.error.message));
  });

  it('fails validation if id is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userId is not a valid guid', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, userId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if userName is not a valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, userName: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if email is not a valid email', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, email: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if only meetingId and id are passed', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if userId is valid', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, userId: userId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if userName is valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, userName: userName });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if email is valid UPN', () => {
    const actual = commandOptionsSchema.safeParse({ meetingId: meetingId, id: attendanceReportId, email: userName });
    assert.strictEqual(actual.success, true);
  });
});