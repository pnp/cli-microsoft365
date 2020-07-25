import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./serviceprincipal-grant-add');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SERVICEPRINCIPAL_GRANT_ADD, () => {
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
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
    cmdInstance.action({ options: { debug: true, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><ObjectPath Id="8" ObjectPathId="7" /><Query Id="9" ObjectPathId="7"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionRequests" /><Method Id="7" ParentId="5" Name="Approve"><Parameters><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></ObjectPaths></Request>`) {
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
    cmdInstance.action({ options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
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
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "A service principal with the name Microsoft Graph1 could not be found.\r\nParameter name: resourceName", "ErrorValue": null, "TraceCorrelationId": "5fdf9d9e-00e9-0000-37ae-14e6d75290f3", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "5fdf9d9e-00e9-0000-37ae-14e6d75290f3"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, resource: 'Microsoft Graph1', scope: 'Mail.Read' } }, (err?: any) => {
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
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "An OAuth permission with the scope Calendar.Read could not be found.\r\nParameter name: permissionRequest", "ErrorValue": null, "TraceCorrelationId": "51df9d9e-d075-0000-37ae-1d4e6902cc2e", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "51df9d9e-d075-0000-37ae-1d4e6902cc2e"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, resource: 'Microsoft Graph', scope: 'Calendar.Read' } }, (err?: any) => {
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
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8224.1210", "ErrorInfo": {
            "ErrorMessage": "An OAuth permission with the resource Microsoft Graph and scope Mail.Read already exists.\r\nParameter name: permissionRequest", "ErrorValue": null, "TraceCorrelationId": "1bdf9d9e-1088-0000-38d6-3b6395428c90", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
          }, "TraceCorrelationId": "1bdf9d9e-1088-0000-38d6-3b6395428c90"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, (err?: any) => {
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
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));
    cmdInstance.action({ options: { debug: false, resource: 'Microsoft Graph', scope: 'Mail.Read' } }, (err?: any) => {
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
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});