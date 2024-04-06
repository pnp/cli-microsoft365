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
import command from './meeting-attendancereport-get.js';
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

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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

  it('retrieves attendance report for currently signed in user', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by userId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, userId: userId, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by userName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, userName: userName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('retrieves attendance report using application permissions by email', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true } });
    assert(loggerLogSpy.calledOnceWith(response));
  });

  it('throws error when using application permissions and not mentioning userId, userName or email', async () => {
    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, verbose: true } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions.`));
  });

  it('throws error when using delegated permissions and mentioning userId, userName or email', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, userId: userId, verbose: true } } as any),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance report using delegated permissions.`));
  });

  it('throws error when meeting not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true } } as any),
      new CommandError(error.error.message));
  });

  it('throws error when attendanceReport not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/onlineMeetings/${meetingId}/attendanceReports/${attendanceReportId}?$expand=attendanceRecords`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, id: attendanceReportId, email: userName, verbose: true } } as any),
      new CommandError(error.error.message));
  });

  it('fails validation if id is not a valid guid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId is not a valid guid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if email is not a valid email', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, email: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only meetingId and id are passed', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userId is valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if userName is valid UPN', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if email is valid UPN', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, id: attendanceReportId, email: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});