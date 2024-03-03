import assert from 'assert';
import sinon, { SinonFakeTimers } from 'sinon';
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
import commands from '../../commands.js';
import command from './meeting-add.js';

describe(commands.MEETING_ADD, () => {
  const startTime = '2022-04-04T03:00:00Z';
  const endTime = '2022-04-04T04:00:00Z';
  const subject = 'test subject';
  const participantUserNames = 'abc@email.com,abc2@email.com';
  const organizerEmail = 'organizer@email.com';

  // #region responses
  const meeting = `{"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#users('1af1bc3a-eb29-4e90-b38a-7ff729d4ac00')/onlineMeetings/$entity","id":"MSoxYWYxYmMzYS1lYjI5LTRlOTAtYjM4YS03ZmY3MjlkNGFjMDAqMCoqMTk6bWVldGluZ19NekkyTnpoak4yTXRaVEk1WmkwMFlUaGpMVGd6WTJNdFl6RTFNamRpTUdSbFpUQTNAdGhyZWFkLnYy","creationDateTime":"2023-10-03T19:13:29.9677485Z","startDateTime":"2023-10-03T19:13:29.596964Z","endDateTime":"2023-10-03T20:13:29.596964Z","joinUrl":"https://teams.microsoft.com/l/meetup-join/19%3ameeting_MzI2N…+4px%3bfont-family%3a%27Segoe+UI%27%2c%27Helvetica+Neue%27%2cHelvetica%2cArial%2csans-serif%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22font-size%3a+12px%3b%22%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%0d%0a%3c%2fdiv%3e%0d%0a%3cdiv+style%3d%22width%3a100%25%3b%22%3e%0d%0a++++%3cspan+style%3d%22white-space%3anowrap%3bcolor%3a%235F5F5F%3bopacity%3a.36%3b%22%3e________________________________________________________________________________%3c%2fspan%3e%0d%0a%3c%2fdiv%3e","contentType":"html"}}`;
  // #endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let fakeTimers: SinonFakeTimers;

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
    fakeTimers = sinon.useFakeTimers(new Date('2020-01-01T12:00:00.000Z'));
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get,
      request.post,
      entraUser.getUserIdByEmail,
      fakeTimers
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MEETING_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('completes validation when no parameters are provided', async () => {
    const actual = await command.validate({ options: { startTime: undefined, endTime: undefined, subject: undefined, participantUserNames: undefined, organizerEmail: undefined, recordAutomatically: undefined } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only the startTime parameter is provided, and it is a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: startTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when both the startTime and endTime parameters are provided, and they are valid ISODateTimes', async () => {
    const actual = await command.validate({ options: { startTime: startTime, endTime: endTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only the subject parameter is provided', async () => {
    const actual = await command.validate({ options: { subject: subject } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only the organizerEmail parameter is provided', async () => {
    const actual = await command.validate({ options: { organizerEmail: organizerEmail } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when only the participantUserNames parameter is provided', async () => {
    const actual = await command.validate({ options: { participantUserNames: participantUserNames } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('completes validation when the correct endTime is provided, and the startTime is not provided', async () => {
    const actual = await command.validate({ options: { endTime: endTime } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when the startTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation startTIme is provided and occurs before the current time.', async () => {
    const actual = await command.validate({ options: { startTime: '1990-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the correct startTime is provided, and the endTime is not a valid ISODateTime', async () => {
    const actual = await command.validate({ options: { startTime: startTime, endTime: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the endTime is before the startTime', async () => {
    const actual = await command.validate({ options: { startTime: '2023-01-01', endTime: '2022-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when only the endTime is provided and occurs before the current time.', async () => {
    const actual = await command.validate({ options: { endTime: '1990-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the organizerEmail parameter is not a valid email address', async () => {
    const actual = await command.validate({ options: { organizerEmail: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participantUserNames is not a valid', async () => {
    const actual = await command.validate({ options: { participantUserNames: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participantUserNames are not separated by comma', async () => {
    const actual = await command.validate({ options: { participantUserNames: 'abc@email.com|abc2@email.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the participantUserNames has incorrect format', async () => {
    const actual = await command.validate({ options: { participantUserNames: 'abc@email.com,foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the startDate is after the endDate', async () => {
    const actual = await command.validate({ options: { startTime: '2023-01-01', endTime: '2022-12-31' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws an error when the organizerEmail is not provided while signed in using app-only authentication', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, { options: { verbose: true } }),
      new CommandError(`The option 'organizerEmail' is required when creating a meeting using app only permissions`));
  });

  it('throws an error when the organizerEmail parameter is set and delegated permissions are used', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
        organizerEmail: organizerEmail
      }
    }), new CommandError(`The option 'organizerEmail' is not supported when creating a meeting using delegated permissions`));
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with a defined startDate for the currently logged-in user', async () => {
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
        endTime: endTime
      }
    });

    assert.strictEqual(calledUrl, 'https://graph.microsoft.com/v1.0/me/onlineMeetings');
    assert.deepEqual(calledData, { startDateTime: '2020-01-01T12:00:00.000Z', endDateTime: endTime });
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with a defined endDate and current date as startDate for the currently logged-in user', async () => {
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with defined startDate and endDate for the currently logged-in user', async () => {
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with defined startDate, endDate and subject for the currently logged-in user', async () => {
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with defined startDate, endDate, subject, and participantUserNames for the currently logged-in user', async () => {
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
        participantUserNames: participantUserNames
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with defined startDate, endDate, subject, participantUserNames and recordAutomatically for the currently logged-in user', async () => {
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
        participantUserNames: participantUserNames,
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('create a meeting with defined startDate, endDate, subject, participantUserNames and recordAutomatically for the specified organizerEmail when app only authorization', async () => {
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
        participantUserNames: participantUserNames,
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
    assert(loggerLogSpy.calledOnceWithExactly(meeting));
  });

  it('handles error appropriately when organizer Id is not found', async () => {
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
        participantUserNames: participantUserNames,
        recordAutomatically: true,
        organizerEmail: organizerEmail
      }
    }), new CommandError(`The specified user with email organizer@email.com does not exist`)
    );
  });

  it('handles error appropriately', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/onlineMeetings') {
        throw {
          "error": {
            "code": "BadRequest",
            "message": "An error has occurred.",
            "innerError": {
              "date": "2023-12-04T20:30:02",
              "request-id": "3fde795d-88a3-45e3-8e22-c428d298c918",
              "client-request-id": "3fde795d-88a3-45e3-8e22-c428d298c918"
            }
          }
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
        participantUserNames: participantUserNames,
        recordAutomatically: true
      }
    }), new CommandError('An error has occurred.'));
  });
});