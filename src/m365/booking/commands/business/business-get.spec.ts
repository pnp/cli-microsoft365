import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./business-get');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.BUSINESS_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'name']]);
  });

  it('defines correct properties for the text output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'businessType', 'phone', 'email', 'defaultCurrencyIso']);
  });

  it('gets business by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return Promise.resolve(businessResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, id: validId } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('gets business by title', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [businessResponse] });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${formatting.encodeQueryParameter(validId)}`) {
        return Promise.resolve(businessResponse);
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: false, name: validName } });
    assert(loggerLogSpy.calledWith(businessResponse));
  });

  it('fails when multiple businesses found with same name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [businessResponse, businessResponse] });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, name: validName } } as any), new CommandError(`Multiple businesses with name ${validName} found. Please disambiguate: ${validId}, ${validId}`));
  });

  it('fails when no business found with name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
  });

  it('fails when no business found with name because of an empty displayName', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [{ 'displayName': null }] });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: false, name: validName } } as any), new CommandError(`The specified business with name ${validName} does not exist.`));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: false } } as any), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});