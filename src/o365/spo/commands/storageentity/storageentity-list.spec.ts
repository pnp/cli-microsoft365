import commands from '../../commands';
import Command, { CommandHelp, CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const storageEntityListCommand: Command = require('./storageentity-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.STORAGEENTITY_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
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
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken
    ]);
  });

  it('has correct name', () => {
    assert.equal(storageEntityListCommand.name.startsWith(commands.STORAGEENTITY_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(storageEntityListCommand.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = storageEntityListCommand.action();
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
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: {}, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert.equal(telemetry.name, commands.STORAGEENTITY_LIST);
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
    cmdInstance.action = storageEntityListCommand.action();
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

  it('retrieves the list of configured tenant properties', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ storageentitiesindex: '' });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: false, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is empty', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({});
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
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
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
            return Promise.resolve({ storageentitiesindex: JSON.stringify({}) });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: false, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
      try {
        assert.equal(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('doesn\'t fail if tenant properties web property value is empty JSON object (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
            return Promise.resolve({ storageentitiesindex: JSON.stringify({}) });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
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
      if (opts.url.indexOf(`/_api/web/AllProperties?$select=storageentitiesindex`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
            return Promise.resolve({ storageentitiesindex: 'a' });
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }}, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Unexpected token a in JSON at position 0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (storageEntityListCommand.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('requires app catalog URL', () => {
    const options = (storageEntityListCommand.options() as CommandOption[]);
    let requiresAppCatalogUrl = false;
    options.forEach(o => {
      if (o.option.indexOf('<appCatalogUrl>') > -1) {
        requiresAppCatalogUrl = true;
      }
    });
    assert(requiresAppCatalogUrl);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return undefined; });
    const options = (storageEntityListCommand.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
  });

  it('accepts valid SharePoint Online app catalog URL', () => {
    const actual = (storageEntityListCommand.validate() as CommandValidate)({ options: { appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }});
    assert(actual);
  });

  it('accepts valid SharePoint Online site URL', () => {
    const actual = (storageEntityListCommand.validate() as CommandValidate)({ options: { appCatalogUrl: 'https://contoso.sharepoint.com' }});
    assert(actual);
  });

  it('rejects invalid SharePoint Online URL', () => {
    const url = 'https://contoso.com';
    const actual = (storageEntityListCommand.validate() as CommandValidate)({ options: { appCatalogUrl: url }});
    assert.equal(actual, `${url} is not a valid SharePoint Online site URL`);
  });

  it('fails validation when no SharePoint Online app catalog URL specified', () => {
    const actual = (storageEntityListCommand.validate() as CommandValidate)({ options: { }});
    assert.equal(actual, 'Missing required option appCatalogUrl');
  });

  it('has help referring to the right command', () => {
    const _helpLog: string[] = [];
    const helpLog = (msg: string) => { _helpLog.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (storageEntityListCommand.help() as CommandHelp)({}, helpLog);
    assert(find.calledWith(commands.STORAGEENTITY_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const log = (msg: string) => { _log.push(msg); }
    const cmd: any = {
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    (storageEntityListCommand.help() as CommandHelp)({}, log);
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
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = storageEntityListCommand.action();
    cmdInstance.action({ options: { debug: true, appCatalogUrl: 'https://contoso-admin.sharepoint.com' }}, () => {
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