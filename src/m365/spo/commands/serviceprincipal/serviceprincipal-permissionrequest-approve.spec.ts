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
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
import * as SpoServicePrincipalPermissionRequestListCommand from './serviceprincipal-permissionrequest-list';
const command: Command = require('./serviceprincipal-permissionrequest-approve');

describe(commands.SERVICEPRINCIPAL_PERMISSIONREQUEST_APPROVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validId = "4dc4c043-25ee-40f2-81d3-b3bf63da7538";

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      telemetry.trackEvent,
      pid.getProcessName
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

  it('approves the specified permission request (debug)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${validId}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
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
    await command.action(logger, { options: { debug: true, id: validId } });
    assert(loggerLogSpy.calledWith({
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "Calendars.ReadWrite"
    }));
  });

  it('approves the specified permission request', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${validId}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
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
    await command.action(logger, { options: { id: validId } });
    assert(loggerLogSpy.calledWith({
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "Calendars.ReadWrite"
    }));
  });

  it('approves all the specified permission request', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoServicePrincipalPermissionRequestListCommand) {
        return ({
          stdout: `[
            {
              "Id": "${validId}",
              "Resource": "Microsoft Graph",
              "ResourceId": "Microsoft Graph",
              "Scope": "Calendars.ReadWrite"
            },
            {
              "Id": "326b80a4-a6e7-43e0-9bb5-893da05e3b72",
              "Resource": "Microsoft Graph",
              "ResourceId": "Microsoft Graph",
              "Scope": "User.Read"
            }
          ]`
        });
      }

      throw new CommandError('Unknown case');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${validId}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
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
      else if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CLI for Microsoft 365 v6.1.0" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{326b80a4-a6e7-43e0-9bb5-893da05e3b72}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>') {
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
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA", "Resource": "Microsoft Graph", "ResourceId": "0e721cc4-302b-4644-bc51-91b41b24d9f0", "Scope": "User.Read"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { all: true } });
    assert(loggerLogSpy.calledWith([{
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "Calendars.ReadWrite"
    }, {
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "User.Read"
    }]));
  });

  it('approves all the permission request by resource', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === SpoServicePrincipalPermissionRequestListCommand) {
        return ({
          stdout: `[
            {
              "Id": "${validId}",
              "Resource": "Microsoft Graph",
              "ResourceId": "Microsoft Graph",
              "Scope": "Calendars.ReadWrite"
            },
            {
              "Id": "326b80a4-a6e7-43e0-9bb5-893da05e3b72",
              "Resource": "Microsoft Graph",
              "ResourceId": "Microsoft Graph",
              "Scope": "User.Read"
            },
            {
              "Id": "9c7d66ae-c9a6-4338-b10b-ad18d0ecf96f",
              "Resource": "Windows Azure Active Directory",
              "ResourceId": "Windows Azure Active Directory",
              "Scope": "User.Read"
            }
          ]`
        });
      }

      throw new CommandError('Unknown case');
    });
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{${validId}}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>`) {
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
      else if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="CLI for Microsoft 365 v6.1.0" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="16" ObjectPathId="15" /><ObjectPath Id="18" ObjectPathId="17" /><ObjectPath Id="20" ObjectPathId="19" /><ObjectPath Id="22" ObjectPathId="21" /><Query Id="23" ObjectPathId="21"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="15" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="17" ParentId="15" Name="PermissionRequests" /><Method Id="19" ParentId="17" Name="GetById"><Parameters><Parameter Type="Guid">{326b80a4-a6e7-43e0-9bb5-893da05e3b72}</Parameter></Parameters></Method><Method Id="21" ParentId="19" Name="Approve" /></ObjectPaths></Request>') {
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
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "cd4043e7-b749-420b-bd07-aa7c3912ed22", "ConsentType": "AllPrincipals", "ObjectId": "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA", "Resource": "Microsoft Graph", "ResourceId": "0e721cc4-302b-4644-bc51-91b41b24d9f0", "Scope": "User.Read"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    await command.action(logger, { options: { resource: "Microsoft Graph" } });
    assert(loggerLogSpy.calledWith([{
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "Calendars.ReadWrite"
    }, {
      ClientId: "cd4043e7-b749-420b-bd07-aa7c3912ed22",
      ConsentType: "AllPrincipals",
      ObjectId: "50NAzUm3C0K9B6p8ORLtIsQccg4rMERGvFGRtBsk2fA",
      Resource: "Microsoft Graph",
      ResourceId: "0e721cc4-302b-4644-bc51-91b41b24d9f0",
      Scope: "User.Read"
    }]));
  });

  it('correctly handles error when approving permission request', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "Permission entry already exists.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "InvalidOperationException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    await assert.rejects(command.action(logger, { options: { id: validId } } as any),
      new CommandError('Permission entry already exists.'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    await assert.rejects(command.action(logger, { options: { id: validId } } as any),
      new CommandError('An error has occurred'));
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('allows specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (id)', async () => {
    const actual = await command.validate({ options: { id: validId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (all)', async () => {
    const actual = await command.validate({ options: { all: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (resource)', async () => {
    const actual = await command.validate({ options: { resource: "Microsoft Graph" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
