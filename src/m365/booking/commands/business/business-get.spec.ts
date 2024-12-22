import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './business-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.BUSINESS_GET, () => {
  const validId = 'mail@contoso.onmicrosoft.com';
  const validName = 'Valid Business';

  const businessResponse = {
    'id': validId,
    'displayName': validName,
    'businessType': 'Other',
    'phone': '',
    'email': 'user@contoso.onmicrosoft.com',
    'webSiteUrl': '',
    'defaultCurrencyIso': 'USD'
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
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
      request.get,
      cli.executeCommandWithOutput,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.BUSINESS_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets business by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return businessResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: validId } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('gets business by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [businessResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return businessResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { name: validName } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('fails when multiple businesses found with same name', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [businessResponse, businessResponse] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError("Multiple businesses with name 'Valid Business' found. Found: mail@contoso.onmicrosoft.com."));
  });

  it('handles selecting single result when multiple businesses with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [businessResponse, businessResponse] });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return Promise.resolve(businessResponse);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return true;
      }

      return defaultValue;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(businessResponse);

    await command.action(logger, { options: { name: validName } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('fails when no business found with name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
  });

  it('fails when no business found with name because of an empty displayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return { value: [{ 'displayName': null }] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
