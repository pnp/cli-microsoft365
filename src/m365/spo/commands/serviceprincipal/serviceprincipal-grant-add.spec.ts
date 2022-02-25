import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./serviceprincipal-grant-add');

describe(commands.SERVICEPRINCIPAL_GRANT_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_GRANT_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('grants the specified API permission (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": null, "TraceCorrelationId": "53df9d9e-50fd-0000-37ae-14a315385835"
          }, 18, {
            "IsNull": false
          }, 20, {
            "IsNull": false
          }, 22, {
            "IsNull": false
          }, 23, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "868668f8-583a-4c66-b3ce-d4e14bc9ceb3", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "-GiGhjpYZkyzztThS8nOs8VG6EHn4S1OjgiedYOfUrQ", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "41e846c5-e1e7-4e2d-8e08-9e75839f52b4", "Scope": "Mail.Read"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "ClientId": "868668f8-583a-4c66-b3ce-d4e14bc9ceb3", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "-GiGhjpYZkyzztThS8nOs8VG6EHn4S1OjgiedYOfUrQ", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "41e846c5-e1e7-4e2d-8e08-9e75839f52b4", "Scope": "Mail.Read"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('grants the specified API permission', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": null, "TraceCorrelationId": "53df9d9e-50fd-0000-37ae-14a315385835"
          }, 18, {
            "IsNull": false
          }, 20, {
            "IsNull": false
          }, 22, {
            "IsNull": false
          }, 23, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "868668f8-583a-4c66-b3ce-d4e14bc9ceb3", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "-GiGhjpYZkyzztThS8nOs8VG6EHn4S1OjgiedYOfUrQ", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "41e846c5-e1e7-4e2d-8e08-9e75839f52b4", "Scope": "Mail.Read"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "ClientId": "868668f8-583a-4c66-b3ce-d4e14bc9ceb3", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "-GiGhjpYZkyzztThS8nOs8VG6EHn4S1OjgiedYOfUrQ", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "41e846c5-e1e7-4e2d-8e08-9e75839f52b4", "Scope": "Mail.Read"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified resource doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "A service principal with the name Microsoft Graph1 could not be found.\r\nParameter name: resourceName", "ErrorValue": null, "TraceCorrelationId": "5fdf9d9e-00e9-0000-37ae-14e6d75290f3", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "5fdf9d9e-00e9-0000-37ae-14e6d75290f3"
        }
      ]));
    });
    command.action(logger, { options: { debug: false, resource: 'Microsoft Graph1', scope: 'Mail.Read' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('A service principal with the name Microsoft Graph1 could not be found.\r\nParameter name: resourceName')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified scope doesn\'t exist', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "An OAuth permission with the scope Calendar.Read could not be found.\r\nParameter name: permissionRequest", "ErrorValue": null, "TraceCorrelationId": "51df9d9e-d075-0000-37ae-1d4e6902cc2e", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "51df9d9e-d075-0000-37ae-1d4e6902cc2e"
        }
      ]));
    });
    command.action(logger, { options: { debug: false, resource: 'Microsoft Graph', scope: 'Calendar.Read' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An OAuth permission with the scope Calendar.Read could not be found.\r\nParameter name: permissionRequest')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified permission has already been granted', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "An OAuth permission with the resource Microsoft Graph and scope Mail.Read already exists.\r\nParameter name: permissionRequest", "ErrorValue": null, "TraceCorrelationId": "1bdf9d9e-1088-0000-38d6-3b6395428c90", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "1bdf9d9e-1088-0000-38d6-3b6395428c90"
        }
      ]));
    });
    command.action(logger, { options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An OAuth permission with the resource Microsoft Graph and scope Mail.Read already exists.\r\nParameter name: permissionRequest')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    command.action(logger, { options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});