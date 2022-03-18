import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError, CommandOption } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./tenant-recyclebinitem-restore');

describe(commands.TENANT_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: futureDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    (command as any).currentContext = undefined;
    sinonUtil.restore([
      request.post,
      global.setTimeout,
      spo.ensureFormDigest
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_RECYCLEBINITEM_RESTORE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { url: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/hr' } });
    assert(actual);
  });

  it('restores the deleted site collection from the tenant recycle bin, doesn\'t wait for completion', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(loggerLogToStderrSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (verbose)', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { verbose: true, url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, doesn\'t wait for completion (debug)', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/hr' } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, wait for completion. Operation immediately completed', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": true
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, wait for completion', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RestoreDeletedSite&#xA;637361531422835228&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, wait for completion (verbose)', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": false
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RestoreDeletedSite&#xA;637361531422835228&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;85db752c-c863-465a-8095-d800361f94b2" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { verbose: true, url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('restores the deleted site collection from the tenant recycle bin, wait for completion (debug)', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 5, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n85db752c-c863-465a-8095-d800361f94b2", "PollingInterval": 15000, "IsComplete": false
            }
          ]));
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f&#xA;SpoOperation&#xA;RestoreDeletedSite&#xA;637361531422835228&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr&#xA;85db752c-c863-465a-8095-d800361f94b2" /></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20509.12004", "ErrorInfo": null, "TraceCorrelationId": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "47ac7b9f-6025-2000-3d94-fb3bb82b6a31|908bed80-a04a-4433-b4a0-883d9847d110:1fdd85e0-9a94-4593-8ab0-5ad1b834475f\nSpoOperation\nRestoreDeletedSite\n637361531422835228\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fhr\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    command.action(logger, { options: { debug: true, url: 'https://contoso.sharepoint.com/sites/hr', wait: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('did not restore the deleted site collection from the tenant recycle bin', (done) => {
    sinonUtil.restore(spo.ensureFormDigest);

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);
    sinon.stub(spo, 'ensureFormDigest').callsFake(() => { return Promise.resolve({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: pastDate, WebFullUrl: 'https://contoso.sharepoint.com/sites/hr' }); });

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties><Property Name="PollingInterval" ScalarProperty="true" /><Property Name="IsComplete" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="3" ParentId="1" Name="RestoreDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/hr</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return Promise.resolve(JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.20308.12005", "ErrorInfo": {
                "ErrorMessage": "Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.", "ErrorValue": null, "TraceCorrelationId": "38ee669f-2049-2000-2f4c-66b5d2e562a2", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "38ee669f-2049-2000-2f4c-66b5d2e562a2"
            }
          ]));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Unable to find the deleted site: https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fhr.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});