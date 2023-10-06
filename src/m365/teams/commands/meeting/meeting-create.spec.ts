import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { aadUser } from '../../../../utils/aadUser.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './meeting-create.js';

describe(commands.MEETING_CREATE, () => {
  const startTime = '2022-04-04T03:00:00Z';
  const endTime = '2022-04-04T04:00:00Z';
  const subject = 'test subject';
  const participants = 'abc@email.com,abc2@email.com';
  const organizerEmail = 'organizer@email.com';

  // #region responses
  const meeting = `{"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#users('1af1bc3a-eb29-4e90-b38a-7ff729d4ac00')/onlineMeetings/$entity","id":"MSoxYWYxYmMzYS1lYjI5LTRlOTAtYjM4YS03ZmY3MjlkNGFjMDAqMCoqMTk6bWVldGluZ19NekkyTnpoak4yTXRaVEk1WmkwMFlUaGpMVGd6WTJNdFl6RTFNamRpTUdSbFpUQTNAdGhyZWFkLnYy","creationDateTime":"2023-10-03T19:13:29.9677485Z","startDateTime":"2023-10-03T19:13:29.596964Z","endDateTime":"2023-10-03T20:13:29.596964Z","joinUrl":"https://teams.microsoft.com/l/meetup-join/19%3ameeting_MzI2N…+4px%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22font-size%3a+12px%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22width%3a100%25%3b%22%3e%0d%0a++++%3cspan+style%3d%22white-space%3anowrap%3bcolor%3a%235F5F5F%3bopacity%3a.36%3b%22%3e________________________________________________________________________________%3c%2fspan%3e%0d%0a%3c%2fdiv%3e","contentType":"html"}}`;

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
    auth.service.connected = true;
    auth.service.accessTokens[auth.defaultResource] = {
      expiresOn: 'abc',
      accessToken: 'abc'
    };
    commandInfo = Cli.getCommandInfo(command);
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
      aadUser.getUserIdByEmail
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_CREATE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('completes validation when no parameter is provided', async () => {
    const actual = await command.validate({ options: { startTime: undefined, endTime: undefined, subject: undefined, participants: undefined, organizerEmail: undefined, recordAutomatically: undefined } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only startTime is provided and it is a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: startTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when the startTime and endTime are provided and they are valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: startTime, endTime: endTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only subject is provided', async () => {
    const actual = await command.validate({ options: { subject: subject } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only organizerEmail is provided', async () => {
    const actual = await command.validate({ options: { organizerEmail: organizerEmail } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only participants parameter is provided', async () => {
    const actual = await command.validate({ options: { participants: participants } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the startTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the correct startTime is provided and the endTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: startTime, endTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the correct endTime is provided and the startTime is not provided', async () => {
    const actual = await command.validate({ options: { endTime: endTime } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });


  it('fails validation when endTime is before startTime', async () => {
    const actual = await command.validate({ options: { startTime: '2023-01-01', endTime: '2022-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the organizerEmail is not a valid', async () => {
    const actual = await command.validate({ options: { organizerEmail: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participants is not a valid', async () => {
    const actual = await command.validate({ options: { participants: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participants are not separated by comma', async () => {
    const actual = await command.validate({ options: { participants: 'abc@email.com|abc2@email.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participants has incorrect email', async () => {
    const actual = await command.validate({ options: { participants: 'abc@email.com,foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when startDateTime is behind endDateTime', async () => {
    const actual = await command.validate({ options: { startDateTime: '2023-01-01', endDateTime: '2022-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws an error when the organizerEmail is not filled in when signed in using app-only authentication', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { verbose: true } }),
      new CommandError(`The option 'organizerEmail' is required when creating a meeting using app only permissions`));
  });

  it('create a meeting for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, {});
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('create a meeting with defined startDate for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, { startDateTime: startTime });
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('create a meeting with defined startDate and endDate for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, { startDateTime: startTime, endDateTime: endTime });
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('create a meeting with defined startDate, endDate and subject for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, { startDateTime: startTime, endDateTime: endTime, subject: subject });
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('create a meeting with defined startDate, endDate, subject, and participants for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, {
      startDateTime: startTime, endDateTime: endTime, subject: subject, participants: {
        attendees: [
          {
            "upn": "abc@email.com"
          },
          {
            "upn": "abc2@email.com"
          }
        ]
      }
    });
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('create a meeting with defined startDate, endDate, subject, participants and recordAutomatically for the currently logged in user', async () => {
    let calledUrl = '';
    let calledData = '';

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants,
        recordAutomatically: true
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, {
      startDateTime: startTime, endDateTime: endTime, subject: subject, participants: {
        attendees: [
          {
            "upn": "abc@email.com"
          },
          {
            "upn": "abc2@email.com"
          }
        ]
      }, recordAutomatically: true
    });
    assert(loggerLogSpy.calledWith(meeting));
  });


  it('create a meeting with defined startDate, endDate, subject, participants and recordAutomatically for the specified organizerEmail when app only authorization', async () => {
    let calledUrl = '';
    let calledData = '';
    const testOrganizerId = '12345678-7882-4efe-b7d1-90703f5a5c65';

    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq 'organizer%40email.com'&$select=id`) {
        return { value: [{ id: testOrganizerId }] };
      }
      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${testOrganizerId}/onlineMeetings`) {
        calledUrl = opts.url;
        calledData = opts.data;
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants,
        recordAutomatically: true,
        organizerEmail: organizerEmail
      }
    });

    assert.strictEqual(calledUrl, `https://graph.microsoft.com/v1.0/users/${testOrganizerId}/onlineMeetings`);
    assert.deepEqual(calledData, {
      startDateTime: startTime, endDateTime: endTime, subject: subject, participants: {
        attendees: [
          {
            "upn": "abc@email.com"
          },
          {
            "upn": "abc2@email.com"
          }
        ]
      }, recordAutomatically: true
    });
    assert(loggerLogSpy.calledWith(meeting));
  });

  it('handles error correctly when organizer Id is not found', async () => {
    const testOrganizerId = '12345678-7882-4efe-b7d1-90703f5a5c65';

    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq 'organizer%40email.com'&$select=id`) {
        return { value: [] };
      }
      throw 'Invalid request: ' + opts.url;
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${testOrganizerId}/onlineMeetings`) {
        return meeting;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants,
        recordAutomatically: true,
        organizerEmail: organizerEmail
      }
    }), new CommandError(`The specified user with email organizer@email.com does not exist`)
    );
  });

  it('handles error forbidden correctly', async () => {

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        throw {
          response: {
            status: 403
          },
          message: 'Forbidden'
        };
      }

      throw 'Invalid request: ' + opts.url;
    });


    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants,
        recordAutomatically: true
      }
    }), new CommandError(`Forbidden. You do not have permission to perform this action. Please verify the command's details for more information.`)
    );
  });

  it('handles error correctly', async () => {

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {

        throw {
          response: {
            status: 404
          },
          message: 'Error message'
        };
      }

      throw 'Invalid request: ' + opts.url;
    });


    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        startTime: startTime,
        endTime: endTime,
        subject: subject,
        participants: participants,
        recordAutomatically: true
      }
    }), new CommandError('Error message')
    );
  });
});