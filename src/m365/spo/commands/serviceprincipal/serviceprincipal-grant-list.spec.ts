import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./serviceprincipal-grant-list');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SERVICEPRINCIPAL_GRANT_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_GRANT_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists permissions granted to the service principal (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionGrants" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "a2a03a9e-602c-4000-879b-1783ec06ba67"
          }, 4, {
            "IsNull": false
          }, 6, {
            "IsNull": false
          }, 7, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrantCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw", "Resource": "Windows Azure Active Directory", "ResourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c", "Scope": "Group.ReadWrite.All"
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg", "Resource": "Microsoft 365 SharePoint Online", "ResourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78", "Scope": "MyFiles.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            ObjectId: '50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw',
            Resource: 'Windows Azure Active Directory',
            ResourceId: '1c444f1a-bba3-42f2-999f-4106c5b1c20c',
            Scope: 'Group.ReadWrite.All'
          },
          {
            ObjectId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg',
            Resource: 'Microsoft 365 SharePoint Online',
            ResourceId: 'dcf25ef3-e2df-4a77-839d-6b7857a11c78',
            Scope: 'MyFiles.Read'
          }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists permissions granted to the service principal', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query><ChildItemQuery SelectAllProperties="true"><Properties /></ChildItemQuery></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="5" ParentId="3" Name="PermissionGrants" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "a2a03a9e-602c-4000-879b-1783ec06ba67"
          }, 4, {
            "IsNull": false
          }, 6, {
            "IsNull": false
          }, 7, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrantCollection", "_Child_Items_": [
              {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw", "Resource": "Windows Azure Active Directory", "ResourceId": "1c444f1a-bba3-42f2-999f-4106c5b1c20c", "Scope": "Group.ReadWrite.All"
              }, {
                "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg", "Resource": "Microsoft 365 SharePoint Online", "ResourceId": "dcf25ef3-e2df-4a77-839d-6b7857a11c78", "Scope": "MyFiles.Read"
              }
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            ObjectId: '50NAzUm3C0K9B6p8ORLtIhpPRByju_JCmZ9BBsWxwgw',
            Resource: 'Windows Azure Active Directory',
            ResourceId: '1c444f1a-bba3-42f2-999f-4106c5b1c20c',
            Scope: 'Group.ReadWrite.All'
          },
          {
            ObjectId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg',
            Resource: 'Microsoft 365 SharePoint Online',
            ResourceId: 'dcf25ef3-e2df-4a77-839d-6b7857a11c78',
            Scope: 'MyFiles.Read'
          }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving permissions granted to the service principal', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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