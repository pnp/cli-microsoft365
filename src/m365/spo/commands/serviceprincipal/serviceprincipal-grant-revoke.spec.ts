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
import command from './serviceprincipal-grant-revoke.js';

describe(commands.SERVICEPRINCIPAL_GRANT_REVOKE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SERVICEPRINCIPAL_GRANT_REVOKE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('revokes the specified permission grant (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
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
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } });
    assert(loggerLogToStderrSpy.called);
  });

  it('revokes the specified permission grant', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return JSON.stringify([
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
            "IsNull": false
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error when revoking permission request', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": {
            "ErrorMessage": "The given key was not present in the dictionary.", "ErrorValue": null, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99", "ErrorCode": -1, "ErrorTypeName": "System.Collections.Generic.KeyNotFoundException"
          }, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99"
        }
      ]);
    });
    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('The given key was not present in the dictionary.'));
  });

  it('revokes the scope from the specified permission grant', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest']) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="3" ParentId="1" Name="PermissionGrants" /><Method Id="5" ParentId="3" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24211.12005", "ErrorInfo": null, "TraceCorrelationId": "a4dee8a0-b086-7000-95cf-1d5b9c2f4276"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "ce0114b0-2529-4227-abc1-5c9157c0bff6", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "7b74a62c-9c27-43bc-b655-737782e64a61", "Scope": "User.Read Mail.Read"
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="Remove" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">ce0114b0-2529-4227-abc1-5c9157c0bff6</Parameter><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="8" ParentId="1" Name="GrantManager" /><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24211.12005", "ErrorInfo": null, "TraceCorrelationId": "a5dee8a0-101e-7000-95cf-1bbee945c389"
            }, 9, {
              "IsNull": false
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg', scope: 'Mail.Read' } });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles error while revoking a scope when the grant is not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest']) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="3" ParentId="1" Name="PermissionGrants" /><Method Id="5" ParentId="3" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": {
                "ErrorMessage": "The given key was not present in the dictionary.", "ErrorValue": null, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99", "ErrorCode": -1, "ErrorTypeName": "System.Collections.Generic.KeyNotFoundException"
              }, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg', scope: 'Mail.Read' } } as any),
      new CommandError('The given key was not present in the dictionary.'));
  });

  it('correctly handles error when revoking a scope fails', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest']) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="2" ObjectPathId="1" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="3" ParentId="1" Name="PermissionGrants" /><Method Id="5" ParentId="3" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24211.12005", "ErrorInfo": null, "TraceCorrelationId": "a4dee8a0-b086-7000-95cf-1d5b9c2f4276"
            }, 2, {
              "IsNull": false
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipalPermissionGrant", "ClientId": "ce0114b0-2529-4227-abc1-5c9157c0bff6", "ConsentType": "AllPrincipals", "IsDomainIsolated": false, "ObjectId": "50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg", "PackageName": null, "Resource": "Microsoft Graph", "ResourceId": "7b74a62c-9c27-43bc-b655-737782e64a61", "Scope": "User.Read Mail.Read"
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="9" ObjectPathId="8" /><Method Name="Remove" Id="10" ObjectPathId="8"><Parameters><Parameter Type="String">ce0114b0-2529-4227-abc1-5c9157c0bff6</Parameter><Parameter Type="String">Microsoft Graph</Parameter><Parameter Type="String">Mail.Read</Parameter></Parameters></Method></Actions><ObjectPaths><Property Id="8" ParentId="1" Name="GrantManager" /><Constructor Id="1" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99", "ErrorCode": -1, "ErrorTypeName": "System.Exception"
              }, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { id: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg', scope: 'Mail.Read' } } as any),
      new CommandError('An error has occurred.'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').callsFake(() => { throw 'An error has occurred'; });
    await assert.rejects(command.action(logger, { options: {} } as any),
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
});
