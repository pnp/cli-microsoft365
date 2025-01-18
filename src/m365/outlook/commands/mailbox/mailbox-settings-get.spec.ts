import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import command from './mailbox-settings-get.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.MAILBOX_SETTINGS_GET, () => {
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MAILBOX_SETTINGS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('retrieves mailbox settings of the signed-in user', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/mailboxSettings') {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      verbose: true
    });

    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWith(mailboxSettingsResponse));
  });

  it('retrieves mailbox settings of a user specified by id', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${ userId }')/mailboxSettings`) {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      userId: userId,
      verbose: true
    });

    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWith(mailboxSettingsResponse));
  });

  it('retrieves mailbox settings of a user specified by user principal name', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users('${userName}')/mailboxSettings`) {
        return mailboxSettingsResponse;
      }

      throw 'Invalid request';
    });

    const result = commandOptionsSchema.safeParse({
      userName: userName,
      verbose: true
    });

    await command.action(logger, {
      options: result.data
    });
    assert(loggerLogSpy.calledOnceWith(mailboxSettingsResponse));
  });

  it('fails retrieve mailbox settings if both userId and userName is specified in app-only mode', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const result = commandOptionsSchema.safeParse({ userId: userId, userName: userName, verbose: true });

    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('When running with application permissions either userId or userName is required, but not both'));
  });

  it('fails retrieve mailbox settings if neither userId nor userName is specified in app-only mode', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    const result = commandOptionsSchema.safeParse({ verbose: true });

    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('When running with application permissions either userId or userName is required'));
  });

  it('fails retrieve mailbox settings of the signed-in user if userId is specified', async () => {
    const result = commandOptionsSchema.safeParse({ userId: userId, verbose: true });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('You can retrieve mailbox settings of other users only if CLI is authenticated in app-only mode'));
  });

  it('fails retrieve mailbox settings of the signed-in user if userName is specified', async () => {
    const result = commandOptionsSchema.safeParse({ userName: userName, verbose: true });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('You can retrieve mailbox settings of other users only if CLI is authenticated in app-only mode'));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });
    const result = commandOptionsSchema.safeParse({ verbose: true });
    await assert.rejects(command.action(logger, { options: result.data }), new CommandError('Invalid request'));
  });
});