import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./storageentity-list');

describe(commands.STORAGEENTITY_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.STORAGEENTITY_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves the list of configured tenant properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({
            storageentitiesindex: JSON.stringify({
              'Property1': {
                Value: 'dolor1'
              },
              'Property2': {
                Comment: 'Lorem2',
                Description: 'ipsum2',
                Value: 'dolor2'
              }
            })
          });
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            Key: 'Property1',
            Description: undefined,
            Comment: undefined,
            Value: 'dolor1'
          },
          {
            Key: 'Property2',
            Description: 'ipsum2',
            Comment: 'Lorem2',
            Value: 'dolor2'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if no tenant properties have been configured', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ storageentitiesindex: '' });
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: false, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is empty', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({});
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let correctResponse: boolean = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No tenant properties found') > -1) {
          correctResponse = true;
        }
      });
      try {
        assert(correctResponse, 'Incorrect response');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ storageentitiesindex: JSON.stringify({}) });
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: false, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ storageentitiesindex: JSON.stringify({}) });
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } }, () => {
      let correctResponse: boolean = false;
      log.forEach(l => {
        if (!l || typeof l !== 'string') {
          return;
        }

        if (l.indexOf('No tenant properties found') > -1) {
          correctResponse = true;
        }
      });
      try {
        assert(correctResponse, 'Incorrect response');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is invalid JSON', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ storageentitiesindex: 'a' });
        }
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unexpected token a in JSON at position 0')));
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

  it('requires app catalog URL', () => {
    const options = command.options();
    let requiresAppCatalogUrl = false;
    options.forEach(o => {
      if (o.option.indexOf('<appCatalogUrl>') > -1) {
        requiresAppCatalogUrl = true;
      }
    });
    assert(requiresAppCatalogUrl);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = command.options();
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts valid SharePoint Online app catalog URL', () => {
    const actual = command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' } });
    assert.strictEqual(actual, true);
  });

  it('accepts valid SharePoint Online site URL', () => {
    const actual = command.validate({ options: { appCatalogUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid SharePoint Online URL', () => {
    const url = 'http://contoso';
    const actual = command.validate({ options: { appCatalogUrl: url } });
    assert.strictEqual(actual, `${url} is not a valid SharePoint Online site URL`);
  });

  it('fails validation when no SharePoint Online app catalog URL specified', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, 'Missing required option appCatalogUrl');
  });

  it('handles promise rejection', (done) => {
    sinon.stub(request, 'get').callsFake(() => Promise.reject('error'));

    command.action(logger, {
      options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }
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