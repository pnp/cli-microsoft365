import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-commsite-enable');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SITE_COMMSITE_ENABLE, () => {
  let log: any[];
  let cmdInstance: any;
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return Promise.resolve('ABC'); });
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
      log: (msg: any) => {
        log.push(msg);
      }
    };
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
    assert.strictEqual(command.name.startsWith(commands.SITE_COMMSITE_ENABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables communication site features on the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables communication site features on the specified site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('escapes XML in user input', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com&gt;</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
          }]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com>' } }, (err?: any) => {
      try {
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles a generic error when setting tenant property', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1 &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><Method Name="EnableCommSite" Id="5" ObjectPathId="3"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Guid">{d604dac3-50d3-405e-9ab9-d4713cda74ef}</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
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

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({ options: { debug: true, url: 'https://contoso.sharepoint.com' } }, (err?: any) => {
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

  it('requires site URL', () => {
    const options = (command.options() as CommandOption[]);
    assert(options.find(o => o.option.indexOf('<url>') > -1));
  });

  it('supports specifying design package ID', () => {
    const options = (command.options() as CommandOption[]);
    assert(options.find(o => o.option.indexOf('[designPackageId]') > -1));
  });

  it('fails validation when no site URL specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {}
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid site URL specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'http://contoso.sharepoint.com' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid site URL specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'https://contoso.sharepoint.com' }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid site URL specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'https://contoso.sharepoint.com' }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation when invalid design package ID specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: 'invalid' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when no design package ID specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'https://contoso.sharepoint.com' }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid design package ID specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { url: 'https://contoso.sharepoint.com', designPackageId: '18eefaa9-ca7b-4ca4-802c-db6d254c533d' }
    });
    assert.strictEqual(actual, true);
  });
});