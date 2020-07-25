import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./theme-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import auth from '../../../../Auth';
import config from '../../../../config';

describe(commands.THEME_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let readFileSyncStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    readFileSyncStub = sinon.stub(fs, 'readFileSync').callsFake(() => '123');
  });

  afterEach(() => {
    Utils.restore([
      fs.readFileSync,
      fs.existsSync,
      fs.lstatSync,
      request.post,
      Utils.isValidTheme
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.THEME_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds theme when correct parameters are passed', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso',
        filePath: 'theme.json',
        isInverted: false
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">{"isInverted":false,"name":"Contoso","palette":123}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`);
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds theme when correct parameters are passed (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        isInverted: true
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="UpdateTenantTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">{"isInverted":true,"name":"Contoso","palette":123}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}"/></ObjectPaths></Request>`);
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false,
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown error command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false,
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if file path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'abc', isInverted: false } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if file path points to a directory', () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'abc', isInverted: false } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if file does not contain a valid theme', () => {
    const stats: fs.Stats = new fs.Stats();
    const theme: string = '{ not valid }';
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    readFileSyncStub.callsFake(() => theme);
    sinon.stub(Utils, 'isValidTheme').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'abc', filePath: 'contoso-blue.json', isInverted: false } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when path points to a valid file', () => {
    const stats: fs.Stats = new fs.Stats();
    const theme = `{
      "themePrimary": "#d81e05",
      "themeLighterAlt": "#fdf5f4",
      "themeLighter": "#f9d6d2",
      "themeLight": "#f4b4ac",
      "themeTertiary": "#e87060",
      "themeSecondary": "#dd351e",
      "themeDarkAlt": "#c31a04",
      "themeDark": "#a51603",
      "themeDarker": "#791002",
      "neutralLighterAlt": "#eeeeee",
      "neutralLighter": "#f5f5f5",
      "neutralLight": "#e1e1e1",
      "neutralQuaternaryAlt": "#d1d1d1",
      "neutralQuaternary": "#c8c8c8",
      "neutralTertiaryAlt": "#c0c0c0",
      "neutralTertiary": "#c2c2c2",
      "neutralSecondary": "#858585",
      "neutralPrimaryAlt": "#4b4b4b",
      "neutralPrimary": "#333333",
      "neutralDark": "#272727",
      "black": "#1d1d1d",
      "white": "#f5f5f5"
    }`;
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);
    readFileSyncStub.callsFake(() => theme);
    sinon.stub(Utils, 'isValidTheme').callsFake(() => true);
    const actual = (command.validate() as CommandValidate)({ options: { name: 'contoso-blue', filePath: 'contoso-blue.json', isInverted: false } });

    assert.strictEqual(actual, true);
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
});