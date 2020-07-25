import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./serviceprincipal-permissionrequest-approve');
import * as assert from 'assert';
import request from '../../../../request';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('approves the specified permission request (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{4dc4c043-25ee-40f2-81d3-b3bf63da7538}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA", "Resource": "Microsoft Graph", "ResourceId": "0e721cc4-302b-4644-bc51-91b41b24d9f0", "Scope": "Calendars.ReadWrite"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
          ConsentType: "AllPrincipals",
          ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
          Resource: "Microsoft Graph",
          ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
          Scope: "Calendars.ReadWrite"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('approves the specified permission request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{4dc4c043-25ee-40f2-81d3-b3bf63da7538}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA", "Resource": "Microsoft Graph", "ResourceId": "0e721cc4-302b-4644-bc51-91b41b24d9f0", "Scope": "Calendars.ReadWrite"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: false, requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
          ConsentType: "AllPrincipals",
          ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
          Resource: "Microsoft Graph",
          ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
          Scope: "Calendars.ReadWrite"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when approving permission request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "Permission entry already exists.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "InvalidOperationException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Permission entry already exists.')));
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

  it('allows specifying requestId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--requestId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the requestId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { requestId: '123' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the requestId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { requestId: '4dc4c043-25ee-40f2-81d3-b3bf63da7538' } });
    assert.strictEqual(actual, true);
  });
});