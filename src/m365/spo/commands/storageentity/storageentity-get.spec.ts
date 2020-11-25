import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./storageentity-get');

describe(commands.STORAGEENTITY_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('existingproperty')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Comment: 'Lorem', Description: 'ipsum', Value: 'dolor' });
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('propertywithoutdescription')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Comment: 'Lorem', Value: 'dolor' });
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('propertywithoutcomments')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Description: 'ipsum', Value: 'dolor' });
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('nonexistingproperty')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "odata.null": true });
        }
      }

      if ((opts.url as string).indexOf(`/_api/web/GetStorageEntity('%23myprop')`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Description: 'ipsum', Value: 'dolor' });
        }
      }

      return Promise.reject('Invalid request');
    });
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

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      request.get,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STORAGEENTITY_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the details of an existing tenant property', (done) => {
    command.action(logger, { options: { debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Key: 'existingproperty',
          Value: 'dolor',
          Description: 'ipsum',
          Comment: 'Lorem'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the details of an existing tenant property without a description', (done) => {
    command.action(logger, { options: { debug: true, key: 'propertywithoutdescription', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Key: 'propertywithoutdescription',
          Value: 'dolor',
          Description: undefined,
          Comment: 'Lorem'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the details of an existing tenant property without a comment', (done) => {
    command.action(logger, { options: { debug: false, key: 'propertywithoutcomments', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Key: 'propertywithoutcomments',
          Value: 'dolor',
          Description: 'ipsum',
          Comment: undefined
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles a non-existent tenant property', (done) => {
    command.action(logger, { options: { debug: false, key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles a non-existent tenant property (debug)', (done) => {
    command.action(logger, { options: { debug: true, key: 'nonexistingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let correctValue: boolean = false;
      log.forEach(l => {
        if (l &&
          typeof l === 'string' &&
          l.indexOf('Property with key nonexistingproperty not found') > -1) {
          correctValue = true;
        }
      });
      try {
        assert(correctValue);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes special characters in property name', (done) => {
    command.action(logger, { options: { debug: true, key: '#myprop', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          Key: '#myprop',
          Value: 'dolor',
          Description: 'ipsum',
          Comment: undefined
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires tenant property name', () => {
    const options = command.options();
    let requiresTenantPropertyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<key>') > -1) {
        requiresTenantPropertyName = true;
      }
    });
    assert(requiresTenantPropertyName);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = command.options();
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('handles promise rejection', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('error'));

    command.action(logger, {
      options: { options: { debug: true, key: '#myprop' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});