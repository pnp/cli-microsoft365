import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './theme-get.js';

describe(commands.THEME_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets theme when correct parameters are passed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: 'Contoso'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">Contoso</Parameter></Parameters></Method></ObjectPaths></Request>`);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].IsInverted, true);
    assert.strictEqual(loggerLogSpy.lastCall.args[0].Name, 'Custom22');
    assert.strictEqual(loggerLogSpy.lastCall.args[0].Palette.themeLight, '#affefe');
  });

  it('gets theme when correct parameters are passed (debug)', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: null, TraceCorrelationId: '6038519e-2044-5000-58c3-114f8e60e920' }, 12, { IsNull: false }, 14, { IsNull: false }, 15, { _ObjectType_: 'Microsoft.Online.SharePoint.TenantManagement.ThemeProperties', IsInverted: true, Name: 'Custom22', Palette: { themeLight: '#affefe' } }]);
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Query Id="15" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="11" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="13" ParentId="11" Name="GetTenantTheme"><Parameters><Parameter Type="String">Contoso</Parameter></Parameters></Method></ObjectPaths></Request>`);
    assert.strictEqual(loggerLogSpy.called, true);
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: { ErrorMessage: 'Unknown Error', ErrorValue: null, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55', ErrorCode: -1, ErrorTypeName: 'Microsoft.SharePoint.Client.UnknownError' }, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55' }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    } as any), new CommandError('Unknown Error'));
  });

  it('handles empty error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ SchemaVersion: '15.0.0.0', LibraryVersion: '16.0.7428.1202', ErrorInfo: { ErrorMessage: '', ErrorValue: null, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55', ErrorCode: -1, ErrorTypeName: 'Microsoft.SharePoint.Client.UnknownError' }, TraceCorrelationId: 'fc38519e-a04a-5000-58c3-143b1f230a55' }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    } as any), new CommandError('ClientSvc unknown error'));
  });

  it('handles request error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        throw 'An error has occurred';
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso'
      }
    } as any), new CommandError('An error has occurred'));
  });
});
