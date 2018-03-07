import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./theme-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.THEME_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let requests: any[];
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
    requests = [];
  });

  afterEach(() => {
    Utils.restore(vorpal.find);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.THEME_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.THEME_GET);
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
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets theme when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso'
      }
    }, () => {

      let correctRequestIssued: boolean = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('gets theme when correct parameters are passed (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]));
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, () => {

      let correctRequestIssued: boolean = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: { ErrorMessage: 'Unknown Error', ErrorValue: null, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55', ErrorCode: -1, ErrorTypeName: 'Microsoft.SharePoint.Client.UnknownError' }, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55' }]));
        }
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, () => {

      requests.forEach(r => {
        if (r.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers['X-RequestDigest'] &&
          r.body) {
        }
      });
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('handles request error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve({ FormDigestValue: 'abc' });
        }
      }

      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.reject('An error has occurred');
        }
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.THEME_GET));
  });

  it('fails validation if name is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: '' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when name is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue' } });
    assert(actual);
  });
});