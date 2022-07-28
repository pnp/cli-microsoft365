import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
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
      appInsights.trackEvent
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

  it('gets business by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${encodeURIComponent(validId)}`) {
        return Promise.resolve(businessResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: validId } }, () => {
      try {
        assert(loggerLogSpy.calledWith(businessResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets business by title', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [businessResponse] });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses/${encodeURIComponent(validId)}`) {
        return Promise.resolve(businessResponse);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: validName } }, () => {
      try {
        assert(loggerLogSpy.calledWith(businessResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple businesses found with same name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [businessResponse, businessResponse] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: validName } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple businesses with name ${validName} found. Please disambiguate: ${validId}, ${validId}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when no business found with name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: validName } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified business with name ${validName} does not exist.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when no business found with name because of an empty displayName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/solutions/bookingBusinesses`) {
        return Promise.resolve({ value: [ { 'displayName': null } ] });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, name: validName } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified business with name ${validName} does not exist.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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