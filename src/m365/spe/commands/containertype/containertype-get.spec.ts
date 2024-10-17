import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command from './containertype-get.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';

describe(commands.CONTAINERTYPE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const containerTypeId = '3ec7c59d-ef31-0752-1ab5-5c343a5e8557';
  const containerTypeName = 'SharePoint Embedded Free Trial Container Type';
  const containerTypedata = {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties",
    "ApplicationRedirectUrl": "",
    "AzureSubscriptionId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "ContainerTypeId": "/Guid(3ec7c59d-ef31-0752-1ab5-5c343a5e8557)/",
    "CreationDate": "5/9/2024",
    "DisplayName": "SharePoint Embedded Free Trial Container Type",
    "ExpiryDate": "6/8/2024",
    "IsBillingProfileRequired": false,
    "OwningAppId": "/Guid(e6b529fe-84db-412a-8e9c-5bbb75f7b037)/",
    "OwningTenantId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "Region": null,
    "ResourceGroup": null,
    "SPContainerTypeBillingClassification": 1
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
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
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError("An error has occurred"));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: containerTypeId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if required options specified (displayName)', async () => {
    const actual = await command.validate({ options: { name: "test container" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves container type by ID', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="49" ObjectPathId="48" /><Method Name="GetSPOContainerTypeById" Id="50" ObjectPathId="48"><Parameters><Parameter Type="Guid">{${containerTypeId}}</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="48" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.25117.12004", "ErrorInfo": null, "TraceCorrelationId": "df0a44a1-c013-9000-9064-f786729ad6a5"
            }
            , 49, {
              "IsNull": false
            }, 50, containerTypedata
          ]);
        }
      }

      throw "Invalid request";
    });

    await command.action(logger, { options: { id: containerTypeId } });
    assert(loggerLogSpy.calledWith(containerTypedata));
  });

  it('retrieves container type by ID (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="49" ObjectPathId="48" /><Method Name="GetSPOContainerTypeById" Id="50" ObjectPathId="48"><Parameters><Parameter Type="Guid">{${containerTypeId}}</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="48" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.25117.12004", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 49, {
              "IsNull": false
            }, 50, containerTypedata
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: containerTypeId, debug: true } });
    assert(loggerLogSpy.calledWith(containerTypedata));
  });

  it('correctly handles error when retrieving container type by ID', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {

          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: containerTypeId, verbose: true } } as any),
      new CommandError("An error has occurred."));
  });

  it('retrieves the container type by name successfully', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 46, {
              "IsNull": false
            }, 47, [containerTypedata]
          ]);
        }
      }

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="49" ObjectPathId="48" /><Method Name="GetSPOContainerTypeById" Id="50" ObjectPathId="48"><Parameters><Parameter Type="Guid">{${containerTypeId}}</Parameter><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="48" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.25117.12004", "ErrorInfo": null, "TraceCorrelationId": "df0a44a1-c013-9000-9064-f786729ad6a5"
            }
            , 49, {
              "IsNull": false
            }, 50, containerTypedata
          ]);
        }
      }
      throw "Invalid request";
    });

    await command.action(logger, { options: { name: containerTypeName } });
    assert(loggerLogSpy.calledWith(containerTypedata));
  });

  it('correctly handles container type not found', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 46, {
              "IsNull": false
            }, 47, []
          ]);
        }
      }
      throw "Invalid request";
    });

    await assert.rejects(command.action(logger, { options: { name: 'test' } } as any), new CommandError("Container type with name 'test' not found"));
  });

});
