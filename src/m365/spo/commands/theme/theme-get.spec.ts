import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./theme-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';
import config from '../../../../config';

describe(commands.THEME_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    Utils.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.THEME_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets theme when correct parameters are passed', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">Contoso</Parameter></Parameters></Method></ObjectPaths></Request>`);
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].IsInverted, true);
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].Name, 'Custom22');
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].Palette.themeLight, '#affefe');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets theme when correct parameters are passed (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">Contoso</Parameter></Parameters></Method></ObjectPaths></Request>`);
        assert.strictEqual(cmdInstanceLogSpy.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: { ErrorMessage: 'Unknown Error', ErrorValue: null, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55', ErrorCode: -1, ErrorTypeName: 'Microsoft.SharePoint.Client.UnknownError' }, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55' }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Unknown Error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles empty error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: { ErrorMessage: '', ErrorValue: null, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55', ErrorCode: -1, ErrorTypeName: 'Microsoft.SharePoint.Client.UnknownError' }, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55' }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
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

  it('handles request error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, (err?: any) => {
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