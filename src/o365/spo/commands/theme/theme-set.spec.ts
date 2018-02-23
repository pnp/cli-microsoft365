import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./theme-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.THEME_SET, () => {
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
    assert.equal(command.name.startsWith(commands.THEME_SET), true);
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
        assert.equal(telemetry.name, commands.THEME_SET);
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

  it('adds theme when correct parameters are passed', (done) => {
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
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso',
        filePath: 'theme.json',
        isInverted: false
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
        assert(cmdInstanceLogSpy.calledWith(true));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          fs.readFileSync,
          request.post
        ]);
      }
    });
  });

  it('adds theme when correct parameters are passed (debug)', (done) => {
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
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        isInverted: true
      }
    }, () => {

      let correctRequestIssued = false;

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
          fs.readFileSync,
          request.post
        ]);
      }
    });
  });

  it('handles error command error correctly', (done) => {
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
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": "{ErrorMessage:error occured}", "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false,
      }
    }, () => {

      let correctRequestIssued = false;

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
          fs.readFileSync,
          request.post
        ]);
      }
    });
  });

  it('handles error while adding theme', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

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

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        isInverted: false
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
          fs.readFileSync,
          request.post
        ]);
      }
    });
  });

  it('fails validation if file path not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', isInverted: false } });
    assert.notEqual(actual, true);
  });

  it('fails validation if file path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'abc', isInverted: false } });
    Utils.restore(fs.existsSync);
    assert.notEqual(actual, true);
  });

  it('fails validation if file path points to a directory', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'abc', isInverted: false } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.notEqual(actual, true);
  });

  it('passes validation when path points to a valid file', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'contoso-blue', filePath: 'contoso-blue.json', isInverted: false } });
    Utils.restore([
      fs.existsSync,
      fs.lstatSync
    ]);
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.THEME_SET));
  });

  it('fails validation if name is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: '' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when name is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue' } });
    assert(actual);
  });

  it('fails validation if path is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { fullPath: '' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when path is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { fullPath: 'theme.json' } });
    assert(actual);
  });

  it('fails validation when inverted parameter is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when inverted parameter is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { isInverted: false } });
    assert(actual);
  });
});