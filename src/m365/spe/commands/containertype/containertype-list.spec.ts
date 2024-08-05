import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './containertype-list.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';

describe(commands.CONTAINERTYPE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const containerTypedata = [{
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties",
    "ApplicationRedirectUrl": null,
    "AzureSubscriptionId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "ContainerTypeId": "/Guid(073269af-f1d2-042d-2ef5-5bdd6ac83115)/",
    "CreationDate": null,
    "DisplayName": "test1",
    "ExpiryDate": null,
    "IsBillingProfileRequired": true,
    "OwningAppId": "/Guid(df4085cc-9a38-4255-badc-5c5225610475)/",
    "OwningTenantId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "Region": null,
    "ResourceGroup": null,
    "SPContainerTypeBillingClassification": 0
  },
  {
    "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties",
    "ApplicationRedirectUrl": null,
    "AzureSubscriptionId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "ContainerTypeId": "/Guid(880ab3bd-5b68-01d4-3744-01a7656cf2ba)/",
    "CreationDate": null,
    "DisplayName": "test1",
    "ExpiryDate": null,
    "IsBillingProfileRequired": true,
    "OwningAppId": "/Guid(50785fde-3082-47ac-a36d-06282ac5c7da)/",
    "OwningTenantId": "/Guid(00000000-0000-0000-0000-000000000000)/",
    "Region": null,
    "ResourceGroup": null,
    "SPContainerTypeBillingClassification": 0
  }];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
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
    assert.strictEqual(command.name, commands.CONTAINERTYPE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['ContainerTypeId', 'DisplayName', 'OwningAppId']);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { debug: true } } as any), new CommandError("An error has occurred"));
  });

  it('retrieves list of Container Type', async () => {
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
            }, 47, containerTypedata
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith(containerTypedata));
  });

  it('retrieves list of Container Type (debug)', async () => {
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
            }, 47, containerTypedata
          ]);
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith(containerTypedata));
  });

  it('correctly handles error when retrieving Container Types', async () => {
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

    await assert.rejects(command.action(logger, { options: { debug: true } } as any),
      new CommandError("An error has occurred."));
  });
});
