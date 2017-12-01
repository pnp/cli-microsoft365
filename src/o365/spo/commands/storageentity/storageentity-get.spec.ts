import commands from '../../commands';
import Command, { CommandHelp, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const storageEntityGetCommand: Command = require('./storageentity-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.STORAGEENTITY_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/GetStorageEntity('existingproperty')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Comment: 'Lorem', Description: 'ipsum', Value: 'dolor' });
        }
      }

      if (opts.url.indexOf(`/_api/web/GetStorageEntity('propertywithoutdescription')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Comment: 'Lorem', Value: 'dolor' });
        }
      }

      if (opts.url.indexOf(`/_api/web/GetStorageEntity('propertywithoutcomments')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ Description: 'ipsum', Value: 'dolor' });
        }
      }

      if (opts.url.indexOf(`/_api/web/GetStorageEntity('nonexistingproperty')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ "odata.null": true });
        }
      }

      if (opts.url.indexOf(`/_api/web/GetStorageEntity('%23myprop')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
            return Promise.resolve({ Description: 'ipsum', Value: 'dolor' });
        }
      }

      return Promise.reject('Invalid request');
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(storageEntityGetCommand.name.startsWith(commands.STORAGEENTITY_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(storageEntityGetCommand.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = storageEntityGetCommand.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = storageEntityGetCommand.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert.equal(telemetry.name, commands.STORAGEENTITY_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = storageEntityGetCommand.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the details of an existing tenant property', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
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
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
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
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
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
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
    cmdInstance.action({ options: { debug: false, key: 'nonexistingproperty' }, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles a non-existent tenant property (debug)', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
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
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
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
    const options = (storageEntityGetCommand.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires tenant property name', () => {
    const options = (storageEntityGetCommand.options() as CommandOption[]);
    let requiresTenantPropertyName = false;
    options.forEach(o => {
      if (o.option.indexOf('<key>') > -1) {
        requiresTenantPropertyName = true;
      }
    });
    assert(requiresTenantPropertyName);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (storageEntityGetCommand.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('has help referring to the right command', () => {
    const _helpLog: string[] = [];
    const helpLog = (msg: string) => { _helpLog.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (storageEntityGetCommand.help() as CommandHelp)({}, helpLog);
    assert(find.calledWith(commands.STORAGEENTITY_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const log = (msg: string) => { _log.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (storageEntityGetCommand.help() as CommandHelp)({}, log);
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityGetCommand.action();
    cmdInstance.action({ options: { debug: true }, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});