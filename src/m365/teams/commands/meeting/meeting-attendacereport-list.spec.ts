import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as userGetCommand from '../../../aad/commands/user/user-get';
const command: Command = require('./meeting-attendancereport-list');

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
  // const meetingResponseText: any = [
  //   {
  //     "subject": "Test",
  //     "start": "2022-06-26T12:30:00.0000000",
  //     "end": "2022-06-26T13:00:00.0000000"
  //   },
  //   {
  //     "subject": "Test",
  //     "start": "2022-04-08T11:30:00.0000000",
  //     "end": "2022-04-08T12:00:00.0000000"
  //   },
  //   {
  //     "subject": "Online meeting test",
  //     "start": "2022-03-15T05:00:00.0000000",
  //     "end": "2022-03-15T05:30:00.0000000"
  //   }
  // ];
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has a correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_ATTENDANCEREPORT_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'totalParticipantCount']);
  });

  it('fails validation when the userId is not a GUID', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('succeeds validation when the userId and meetingId are valid', async () => {
    const actual = await command.validate({ options: { meetingId: meetingId, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
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
      {
        meetingId: meetingId
      }
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
      {
        meetingId: meetingId,
        userId: userId
      }
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userName', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/onlineMeetings/${meetingId}/attendanceReports`) {
        return { value: meetingAttendanceResponse };
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

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('retrieves meeting attendace reports correctly by userEmail', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === userGetCommand) {
        return { "stdout": JSON.stringify({ id: userId }) };
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
      {
        meetingId: meetingId,
        email: userName,
        verbose: true
      }
    });

    assert(loggerLogSpy.calledWith(meetingAttendanceResponse));
  });

  it('correctly handles error when throwing request', async () => {
    sinon.stub(request, 'get').callsFake(async () => {
      throw 'An error has occurred';
    });

    await assert.rejects(command.action(logger, { options: { verbose: true, meetingId: meetingId } } as any),
      new CommandError('An error has occurred'));
  });

  it('correctly handles error when options are missing', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meeting attendance report using app only permissions`));
  });

  it('correctly handles error when options are missing with a delegated token', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, { options: { meetingId: meetingId, userId: userId } } as any),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meeting attendance reports using delegated permissions`));
  });
});