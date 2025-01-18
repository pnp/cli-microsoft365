import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './mailbox-settings-set.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.MAILBOX_SETTINGS_SET, () => {
  const userId = 'abcd1234-de71-4623-b4af-96380a352509';
  const userName = 'john.doe@contoso.com';

  const mailboxSettingsResponse = {
    "timeZone": "Central Europe Standard Time",
    "delegateMeetingMessageDeliveryOptions": "sendToDelegateAndInformationToPrincipal",
    "dateFormat": "dd.MM.yyyy",
    "timeFormat": "HH:mm",
    "userPurpose": "user",
    "automaticRepliesSetting": {
      "status": "disabled",
      "externalAudience": "none",
      "internalReplyMessage": "I'm out of office. Contact my manager in case of any troubles.﻿",
      "externalReplyMessage": "I'm out of office",
      "scheduledStartDateTime": {
        "dateTime": "2025-01-03T19:00:00.0000000",
        "timeZone": "UTC"
      },
      "scheduledEndDateTime": {
        "dateTime": "2025-01-04T19:00:00.0000000",
        "timeZone": "UTC"
      }
    },
    "language": {
      "locale": "cs-CZ",
      "displayName": "Czech (Czech Republic)"
    },
    "workingHours": {
      "daysOfWeek": [
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday"
      ],
      "startTime": "07:00:00.0000000",
      "endTime": "16:00:00.0000000",
      "timeZone": {
        "name": "Central Europe Standard Time"
      }
    }
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;
  let commandOptionsSchema: z.ZodTypeAny;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    if (!auth.connection.accessTokens[auth.defaultResource]) {
      auth.connection.accessTokens[auth.defaultResource] = {
        expiresOn: 'abc',
        accessToken: 'abc'
      };
    }
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MAILBOX_SETTINGS_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both userId and userName are specified', () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is provided in delegated mode', () => {
    const actual = commandOptionsSchema.safeParse({ userId: userId });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is provided in delegated mode', () => {
    const actual = commandOptionsSchema.safeParse({ userName: userName });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if both userId and userName are provided in delegated mode', () => {
    const actual = commandOptionsSchema.safeParse({
      userId: userId,
      userName: userName
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userId is not a valid GUID', () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const actual = commandOptionsSchema.safeParse({
      userId: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if userName is not a valid UPN', () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const actual = commandOptionsSchema.safeParse({
      userName: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no option except user id provided in app-only mode', () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const actual = commandOptionsSchema.safeParse({
      userId: userId
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if no option is provided in delegated mode', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if delegateMeetingMessageDeliveryOptions has wrong value', () => {
    const actual = commandOptionsSchema.safeParse({
      delegateMeetingMessageDeliveryOptions: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if workingDays has wrong value', () => {
    const actual = commandOptionsSchema.safeParse({
      workingDays: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if autoReplyExternalAudience has wrong value', () => {
    const actual = commandOptionsSchema.safeParse({
      autoReplyExternalAudience: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation if autoReplyStatus has wrong value', () => {
    const actual = commandOptionsSchema.safeParse({
      autoReplyStatus: 'foo'
    });
    assert.notStrictEqual(actual.success, true);
  });

  it('updates mailbox settings of the signed-in user', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      dateFormat: 'dd.MM.yyy',
      timeFormat: 'HH:mm',
      timeZone: 'Central Europe Standard Time',
      language: 'en-US',
      delegateMeetingMessageDeliveryOptions: 'sendToDelegateAndInformationToPrincipal',
      workingDays: 'monday,tuesday,wednesday,thursday,friday',
      workingHoursStartTime: '09:00:00.000000',
      workingHoursEndTime: '17:00:00.000000',
      workingHoursTimeZone: 'UTC',
      autoReplyExternalAudience: 'contactsOnly',
      autoReplyExternalMessage: "I'm out of office",
      autoReplyInternalMessage: "I'm out of office. Contact my manager in case of any troubles.",
      autoReplyStartDateTime: '2025-01-06T00:00:00.0000000',
      autoReplyStartTimeZone: 'UTC',
      autoReplyEndDateTime: '2025-01-10T00:00:00.0000000',
      autoReplyEndTimeZone: 'UTC',
      autoReplyStatus: 'scheduled',
      verbose: true
    });

    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWith(mailboxSettingsResponse));
  });

  it('updates working hours of a user specified by id', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userId}')/mailboxSettings`) {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      userId: userId,
      workingDays: 'monday,tuesday,wednesday,thursday,friday',
      workingHoursStartTime: '09:00:00.000000',
      workingHoursEndTime: '17:00:00.000000',
      workingHoursTimeZone: 'UTC',
      verbose: true
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      workingHours: {
        daysOfWeek: [
          "monday",
          "tuesday",
          "wednesday",
          "thursday",
          "friday"
        ],
        startTime: "09:00:00.000000",
        endTime: "17:00:00.000000",
        timeZone: {
          name: "UTC"
        }
      }
    });
  });

  it('updates working days of signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/mailboxSettings`) {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      workingDays: 'monday,tuesday,wednesday,thursday,friday'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      workingHours: {
        daysOfWeek: [
          "monday",
          "tuesday",
          "wednesday",
          "thursday",
          "friday"
        ]
      }
    });
  });

  it('updates working hours timezone of a user specified by UPN', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/mailboxSettings`) {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      userName: userName,
      workingHoursTimeZone: 'UTC',
      verbose: true
    });

    await command.action(logger, { options: result.data });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      workingHours: {
        timeZone: {
          name: "UTC"
        }
      }
    });
  });

  it('updates working hours start time of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      workingHoursStartTime: '07:00:00.0000000'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      workingHours: {
        startTime: '07:00:00.0000000'
      }
    });
  });

  it('updates working hours end time of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      workingHoursEndTime: '16:00:00.0000000'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      workingHours: {
        endTime: '16:00:00.0000000'
      }
    });
  });

  it('updates auto reply external audience of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyExternalAudience: 'all'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        externalAudience: 'all'
      }
    });
  });

  it('updates auto reply external message of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyExternalMessage: `I'm out of office`
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        externalReplyMessage: `I'm out of office`
      }
    });
  });

  it('updates auto reply internal message of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyInternalMessage: `I'm out of office. Contact my manager in case of any troubles.`
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        internalReplyMessage: `I'm out of office. Contact my manager in case of any troubles.`
      }
    });
  });

  it('updates auto reply start date time of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyStartDateTime: '2025-01-03T19:00:00.0000000'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        scheduledStartDateTime: {
          dateTime: '2025-01-03T19:00:00.0000000'
        }
      }
    });
  });

  it('updates auto reply start time zone of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyStartTimeZone: 'UTC'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        scheduledStartDateTime: {
          timeZone: 'UTC'
        }
      }
    });
  });

  it('updates auto reply end date time of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyEndDateTime: '2025-01-04T19:00:00.0000000'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        scheduledEndDateTime: {
          dateTime: '2025-01-04T19:00:00.0000000'
        }
      }
    });
  });

  it('updates auto reply end time zone of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyEndTimeZone: 'UTC'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        scheduledEndDateTime: {
          timeZone: 'UTC'
        }
      }
    });
  });

  it('updates auto reply status of the signed-in user', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      autoReplyStatus: 'scheduled'
    });

    await command.action(logger, {
      options: result.data
    });

    assert.deepStrictEqual(patchStub.lastCall.args[0].data, {
      automaticRepliesSetting: {
        status: 'scheduled'
      }
    });
  });

  it('fails updating mailbox settings if neither userId nor userName is specified in app-only mode', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const result = commandOptionsSchema.safeParse({ timeFormat: 'HH:mm', verbose: true });

    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('When running with application permissions either userId or userName is required'));
  });

  it('fails updating mailbox settings of the signed-in user if userId is specified', async () => {
    const result = commandOptionsSchema.safeParse({ userId: userId, timeFormat: 'HH:mm', verbose: true });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('You can update mailbox settings of other users only if CLI is authenticated in app-only mode'));
  });

  it('fails updating mailbox settings of the signed-in user if userName is specified', async () => {
    const result = commandOptionsSchema.safeParse({ userName: userName, timeFormat: 'HH:mm', verbose: true });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('You can update mailbox settings of other users only if CLI is authenticated in app-only mode'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });
    const result = commandOptionsSchema.safeParse({ dateFormat: 'dd.MM.yyy' });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('Invalid request'));
  });
});