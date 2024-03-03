import assert from 'assert';
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
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { formatting } from '../../../../utils/formatting.js';
import commands from '../../commands.js';
import command from './meeting-list.js';

describe(commands.MEETING_LIST, () => {
  const userId = '68be84bf-a585-4776-80b3-30aa5207aa21';
  const startDateTime = '2022-01-01';
  const endDateTime = '2022-12-31';
  const userName = 'user@tenant.com';

  // #region responses
  const meetings = [
    {
      "id": "MSpiMjA5MWUxOC03ODgyLTRlZmUtYjdkMS05MDcwM2Y1YTVjNjUqMCoqMTk6bWVldGluZ19NakEyWkRrNU5tSXRZak15TVMwMFpURTVMVGxqWW1ZdE9ERmpaVGhrTkRVd016ZGlAdGhyZWFkLnYy",
      "creationDateTime": "2023-07-25T19:29:32.033109Z",
      "startDateTime": "2023-07-17T03:00:00Z",
      "endDateTime": "2023-07-17T04:00:00Z",
      "joinUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MjA2ZDk5NmItYjMyMS00ZTE5LTljYmYtODFjZThkNDUwMzdi%40thread.v2/0?context=%7b%22Tid%22%3a%22ad4f158a-97c7-4914-a9bd-038ecde40ff3%22%2c%22Oid%22%3a%22b2091e18-7882-4efe-b7d1-90703f5a5c65%22%7d",
      "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/19%3ameeting_MjA2ZDk5NmItYjMyMS00ZTE5LTljYmYtODFjZThkNDUwMzdi%40thread.v2/0?context=%7b%22Tid%22%3a%22ad4f158a-97c7-4914-a9bd-038ecde40ff3%22%2c%22Oid%22%3a%22b2091e18-7882-4efe-b7d1-90703f5a5c65%22%7d",
      "meetingCode": "396464591835",
      "subject": "Team meeting",
      "isBroadcast": false,
      "autoAdmittedUsers": "unknownFutureValue",
      "outerMeetingAutoAdmittedUsers": null,
      "isEntryExitAnnounced": false,
      "allowedPresenters": "everyone",
      "allowMeetingChat": "enabled",
      "shareMeetingChatHistoryDefault": "none",
      "allowTeamworkReactions": true,
      "allowAttendeeToEnableMic": true,
      "allowAttendeeToEnableCamera": true,
      "recordAutomatically": false,
      "anonymizeIdentityForRoles": [],
      "capabilities": [],
      "videoTeleconferenceId": null,
      "externalId": null,
      "iCalUid": null,
      "meetingType": null,
      "allowParticipantsToChangeName": false,
      "allowRecording": true,
      "allowTranscription": true,
      "meetingMigrationMode": null,
      "broadcastSettings": null,
      "audioConferencing": null,
      "meetingInfo": null,
      "participants": {
        "organizer": {
          "upn": "john.doe@contoso.com",
          "role": "presenter",
          "identity": {
            "application": null,
            "device": null,
            "user": {
              "id": "b2091e18-7882-4efe-b7d1-90703f5a5c65",
              "displayName": null,
              "tenantId": "ad4f158a-97c7-4914-a9bd-038ecde40ff3",
              "identityProvider": "AAD"
            }
          }
        },
        "attendees": [
          {
            "upn": "adele.vance@contoso.com",
            "role": "attendee",
            "identity": {
              "application": null,
              "device": null,
              "user": {
                "id": "52bd2d9c-2d89-416f-96c4-ca94245e22c8",
                "displayName": null,
                "tenantId": "ad4f158a-97c7-4914-a9bd-038ecde40ff3",
                "identityProvider": "AAD"
              }
            }
          }
        ]
      },
      "lobbyBypassSettings": {
        "scope": "unknownFutureValue",
        "isDialInBypassEnabled": false
      },
      "joinMeetingIdSettings": {
        "isPasscodeRequired": true,
        "joinMeetingId": "396464591835",
        "passcode": "Z3GYtQ"
      },
      "chatInfo": {
        "threadId": "19:meeting_MjA2ZDk5NmItYjMyMS00ZTE5LTljYmYtODFjZThkNDUwMzdi@thread.v2",
        "messageId": "0",
        "replyChainMessageId": null
      },
      "joinInformation": {
        "content": "data:text/html,%3cdiv+style%3d%22width%3a100%25%3b%22%3e%0d%0a++++%3cspan+style%3d%22white-space%3anowrap%3bcolor%3a%235F5F5F%3bopacity%3a.36%3b%22%3e________________________________________________________________________________%3c%2fspan%3e%0d%0a%3c%2fdiv%3e%0d%0a+%0d%0a+%3cdiv+class%3d%22me-email-text%22+style%3d%22color%3a%23252424%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22+lang%3d%22en-US%22%3e%0d%0a++++%3cdiv+style%3d%22margin-top%3a+24px%3b+margin-bottom%3a+20px%3b%22%3e%0d%0a++++++++%3cspan+style%3d%22font-size%3a+24px%3b+color%3a%23252424%22%3eMicrosoft+Teams+meeting%3c%2fspan%3e%0d%0a++++%3c%2fdiv%3e%0d%0a++++%3cdiv+style%3d%22margin-bottom%3a+20px%3b%22%3e%0d%0a++++++++%3cdiv+style%3d%22margin-top%3a+0px%3b+margin-bottom%3a+0px%3b+font-weight%3a+bold%22%3e%0d%0a++++++++++%3cspan+style%3d%22font-size%3a+14px%3b+color%3a%23252424%22%3eJoin+on+your+computer%2c+mobile+app+or+room+device%3c%2fspan%3e%0d%0a++++++++%3c%2fdiv%3e%0d%0a++++++++%3ca+class%3d%22me-email-headline%22+style%3d%22font-size%3a+14px%3bfont-family%3a%27Segoe+UI+Semibold%27%2c%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3btext-decoration%3a+underline%3bcolor%3a+%236264a7%3b%22+href%3d%22https%3a%2f%2fteams.microsoft.com%2fl%2fmeetup-join%2f19%253ameeting_MjA2ZDk5NmItYjMyMS00ZTE5LTljYmYtODFjZThkNDUwMzdi%2540thread.v2%2f0%3fcontext%3d%257b%2522Tid%2522%253a%2522ad4f158a-97c7-4914-a9bd-038ecde40ff3%2522%252c%2522Oid%2522%253a%2522b2091e18-7882-4efe-b7d1-90703f5a5c65%2522%257d%22+target%3d%22_blank%22+rel%3d%22noreferrer+noopener%22%3eClick+here+to+join+the+meeting%3c%2fa%3e%0d%0a++++%3c%2fdiv%3e%0d%0a++++%3cdiv+style%3d%22margin-bottom%3a20px%3b+margin-top%3a20px%22%3e%0d%0a++++%3cdiv+style%3d%22margin-bottom%3a4px%22%3e%0d%0a++++++++%3cspan+data-tid%3d%22meeting-code%22+style%3d%22font-size%3a+14px%3b+color%3a%23252424%3b%22%3e%0d%0a++++++++++++Meeting+ID%3a+%3cspan+style%3d%22font-size%3a16px%3b+color%3a%23252424%3b%22%3e396+464+591+835%3c%2fspan%3e%0d%0a+++++++%3c%2fspan%3e%0d%0a+++++++++++%3cbr+%2f%3e%3cspan+style%3d%22font-size%3a+14px%3b+color%3a%23252424%3b%22%3e+Passcode%3a+%3c%2fspan%3e+%3cspan+style%3d%22font-size%3a+16px%3b+color%3a%23252424%3b%22%3e+Z3GYtQ+%3c%2fspan%3e%0d%0a++++++++%3cdiv+style%3d%22font-size%3a+14px%3b%22%3e%3ca+class%3d%22me-email-link%22+style%3d%22font-size%3a+14px%3btext-decoration%3a+underline%3bcolor%3a+%236264a7%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22+target%3d%22_blank%22+href%3d%22https%3a%2f%2fwww.microsoft.com%2fen-us%2fmicrosoft-teams%2fdownload-app%22+rel%3d%22noreferrer+noopener%22%3e%0d%0a++++++++Download+Teams%3c%2fa%3e+%7c+%3ca+class%3d%22me-email-link%22+style%3d%22font-size%3a+14px%3btext-decoration%3a+underline%3bcolor%3a+%236264a7%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22+target%3d%22_blank%22+href%3d%22https%3a%2f%2fwww.microsoft.com%2fmicrosoft-teams%2fjoin-a-meeting%22+rel%3d%22noreferrer+noopener%22%3eJoin+on+the+web%3c%2fa%3e%3c%2fdiv%3e%0d%0a++++%3c%2fdiv%3e%0d%0a+%3c%2fdiv%3e%0d%0a++++%0d%0a++++++%0d%0a++++%0d%0a++++%0d%0a++++%0d%0a++++%3cdiv+style%3d%22margin-bottom%3a+24px%3bmargin-top%3a+20px%3b%22%3e%0d%0a++++++++%3ca+class%3d%22me-email-link%22+style%3d%22font-size%3a+14px%3btext-decoration%3a+underline%3bcolor%3a+%236264a7%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22+target%3d%22_blank%22+href%3d%22https%3a%2f%2faka.ms%2fJoinTeamsMeeting%22+rel%3d%22noreferrer+noopener%22%3eLearn+More%3c%2fa%3e++%7c+%3ca+class%3d%22me-email-link%22+style%3d%22font-size%3a+14px%3btext-decoration%3a+underline%3bcolor%3a+%236264a7%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22+target%3d%22_blank%22+href%3d%22https%3a%2f%2fteams.microsoft.com%2fmeetingOptions%2f%3forganizerId%3db2091e18-7882-4efe-b7d1-90703f5a5c65%26tenantId%3dad4f158a-97c7-4914-a9bd-038ecde40ff3%26threadId%3d19_meeting_MjA2ZDk5NmItYjMyMS00ZTE5LTljYmYtODFjZThkNDUwMzdi%40thread.v2%26messageId%3d0%26language%3den-US%22+rel%3d%22noreferrer+noopener%22%3eMeeting+options%3c%2fa%3e+%0d%0a++++++%3c%2fdiv%3e%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22font-size%3a+14px%3b+margin-bottom%3a+4px%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22font-size%3a+12px%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22width%3a100%25%3b%22%3e%0d%0a++++%3cspan+style%3d%22white-space%3anowrap%3bcolor%3a%235F5F5F%3bopacity%3a.36%3b%22%3e________________________________________________________________________________%3c%2fspan%3e%0d%0a%3c%2fdiv%3e",
        "contentType": "html"
      },
      "watermarkProtection": {
        "isEnabledForContentSharing": false,
        "isEnabledForVideo": false
      }
    }
  ];

  const calendarMeetingsResponse = {
    value: [
      {
        onlineMeeting: {
          joinUrl: meetings[0].joinWebUrl
        }
      }
    ]
  };

  const graphBatchResponse = {
    responses: [
      {
        id: '1',
        status: 200,
        body: {
          value: meetings
        }
      }
    ]
  };

  // #endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      request.post,
      entraUser.getUserIdByEmail
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['subject', 'startDateTime', 'endDateTime']);
  });

  it('completes validation when the startDateTime is a valid ISODateTime, endDateTime is a valid ISODateTime and userId is a valid Guid', async () => {
    const actual = await command.validate({ options: { startDateTime: startDateTime, endDateTime: endDateTime, userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the startDateTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startDateTime: 'foo', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userId is not a valid guid', async () => {
    const actual = await command.validate({ options: { startDateTime: startDateTime, userId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { startDateTime: startDateTime, userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the email is not a valid UPN', async () => {
    const actual = await command.validate({ options: { startDateTime: startDateTime, email: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when startDateTime is behind endDateTime', async () => {
    const actual = await command.validate({ options: { startDateTime: '2023-01-01', endDateTime: '2022-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the endDateTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startDateTime: startDateTime, endDateTime: 'foo', userId: userId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws an error when the userName, userId or email is not filled in when signed in using app-only authentication', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { startDateTime: '2022-04-04' } } as any),
      new CommandError(`The option 'userId', 'userName' or 'email' is required when retrieving meetings using app only permissions`));
  });

  it('throws an error when the userName is filled in when signed in using delegated authentication', async () => {
    await assert.rejects(command.action(logger, { options: { startDateTime: '2022-04-04', email: userName } } as any),
      new CommandError(`The options 'userId', 'userName' and 'email' cannot be used when retrieving meetings using delegated permissions`));
  });

  it('logs meetings for the currently logged in user', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${startDateTime}' and end/dateTime lt '${endDateTime}' and isOrganizer eq true&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        endDateTime: endDateTime,
        isOrganizer: true
      }
    });

    assert(loggerLogSpy.calledWith(meetings));
  });

  it('logs meetings for a user specified by userId', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        userId: userId
      }
    });

    assert(loggerLogSpy.calledWith(meetings));
  });

  it('logs meetings for a user specified by userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        userName: userName
      }
    });

    assert(loggerLogSpy.calledWith(meetings));
  });

  it('logs meetings for a user specified by email', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(entraUser, 'getUserIdByEmail').resolves(userId);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        email: userName
      }
    });

    assert(loggerLogSpy.calledWith(meetings));
  });

  it('filters out non meeting events when retrieving calendar events', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return {
          value: [
            calendarMeetingsResponse.value[0],
            {
              onlineMeeting: null
            }
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime
      }
    });

    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      requests: [
        {
          id: 0,
          method: 'GET',
          url: `me/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(calendarMeetingsResponse.value[0].onlineMeeting.joinUrl)}'`
        }
      ]
    });
  });

  it('retrieves meetings correctly when specifying userId', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        userId: userId
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      requests: [
        {
          id: 0,
          method: 'GET',
          url: `users/${userId}/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(calendarMeetingsResponse.value[0].onlineMeeting.joinUrl)}'`
        }
      ]
    });
  });

  it('retrieves meetings correctly when specifying userName', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userName}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        userName: userName
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      requests: [
        {
          id: 0,
          method: 'GET',
          url: `users/${userName}/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(calendarMeetingsResponse.value[0].onlineMeeting.joinUrl)}'`
        }
      ]
    });
  });

  it('retrieves meetings correctly when specifying email', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(entraUser, 'getUserIdByEmail').resolves(userId);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${userId}/events?$filter=start/dateTime ge '${startDateTime}'&$select=onlineMeeting`) {
        return calendarMeetingsResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return graphBatchResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        email: userName
      }
    });

    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      requests: [
        {
          id: 0,
          method: 'GET',
          url: `users/${userId}/onlineMeetings?$filter=joinWebUrl eq '${formatting.encodeQueryParameter(calendarMeetingsResponse.value[0].onlineMeeting.joinUrl)}'`
        }
      ]
    });
  });

  it('handles error correctly when retrieving calendar events', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').rejects({
      error: {
        error: {
          message: 'User could not be found.'
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime,
        userId: userId
      }
    }), new CommandError('User could not be found.'));
  });

  it('handles error correctly when retrieving meetings', async () => {
    sinon.stub(request, 'get').resolves(calendarMeetingsResponse);

    sinon.stub(request, 'post').resolves({
      responses: [
        {
          id: '0',
          status: 404,
          body: {
            error: {
              message: 'Something went wrong.'
            }
          }
        }
      ]
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime
      }
    }), new CommandError('Something went wrong.'));
  });

  it('handles error without message correctly when retrieving meetings', async () => {
    sinon.stub(request, 'get').resolves(calendarMeetingsResponse);

    sinon.stub(request, 'post').resolves({
      responses: [
        {
          id: '0',
          status: 404,
          body: {
            error: {
              message: '',
              code: 'Forbidden'
            }
          }
        }
      ]
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startDateTime: startDateTime
      }
    }), new CommandError('Forbidden'));
  });
});
