import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./storageentity-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.STORAGEENTITY_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    cmdInstance.action({ options: { debug: true, key: 'existingproperty' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    cmdInstance.action({ options: { debug: true, key: 'propertywithoutdescription' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    cmdInstance.action({ options: { debug: false, key: 'propertywithoutcomments' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    cmdInstance.action({ options: { debug: false, key: 'nonexistingproperty' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
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
    cmdInstance.action({ options: { debug: true, key: 'nonexistingproperty' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
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
    cmdInstance.action({ options: { debug: true, key: '#myprop' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires tenant property name', () => {
    const options = (command.options() as CommandOption[]);
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
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('handles promise rejection', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('error'));

    cmdInstance.action({
      options: { options: { debug: true, key: '#myprop' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }
    }, (err?: any) => {
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