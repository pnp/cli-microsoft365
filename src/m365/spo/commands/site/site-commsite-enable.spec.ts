import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./site-commsite-enable');

describe(commands.SITE_COMMSITE_ENABLE, () => {
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => { return defaultValue; }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_COMMSITE_ENABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables communication site features on the specified site', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } } as any);
  });

  it('enables communication site features on the specified site (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com' } } as any);
  });

  it('escapes XML in user input', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com&gt;</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com>' } } as any);
  });

  it('correctly handles a generic error when setting tenant property', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.PublicCdn.TenantCdnAdministrationException"
            }, "TraceCorrelationId": "965d299e-a0c6-4000-8546-cc244881a129"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com' } } as any), new CommandError('An error has occurred'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com' } } as any), new CommandError('An error has occurred'));
  });

  it('requires site URL', () => {
    const options = command.options;
    assert(options.find(o => o.option.indexOf('<url>') > -1));
  });

  it('supports specifying design package ID', () => {
    const options = command.options;
    assert(options.find(o => o.option.indexOf('[designPackageId]') > -1));
  });

  it('fails validation when no site URL specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid site URL specified', async () => {
    const actual = await command.validate({
      options: { url: 'http://contoso.sharepoint.com' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid site URL specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when invalid design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: 'invalid' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid design package ID specified', async () => {
    const actual = await command.validate({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: '18eefaa9-ca7b-4ca4-802c-db6d254c533d' }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
