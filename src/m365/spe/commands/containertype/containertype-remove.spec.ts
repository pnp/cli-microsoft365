import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './containertype-remove.js';

describe(commands.CONTAINERTYPE_REMOVE, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let confirmationPromptStub: sinon.SinonStub;

  const CsomContainerTypeResponse = [
    {
      _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
      ApplicationRedirectUrl: null,
      AzureSubscriptionId: '/Guid(00000000-0000-0000-0000-000000000000)/',
      ContainerTypeId: `/Guid(${containerTypeId})/`,
      CreationDate: null,
      DisplayName: containerTypeName,
      ExpiryDate: null,
      IsBillingProfileRequired: true,
      OwningAppId: '/Guid(df4085cc-9a38-4255-badc-5c5225610475)/',
      OwningTenantId: '/Guid(00000000-0000-0000-0000-000000000000)/',
      Region: null,
      ResourceGroup: null,
      SPContainerTypeBillingClassification: 0
    },
    {
      _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
      ApplicationRedirectUrl: null,
      AzureSubscriptionId: '/Guid(00000000-0000-0000-0000-000000000000)/',
      ContainerTypeId: '/Guid(880ab3bd-5b68-01d4-3744-01a7656cf2ba)/',
      CreationDate: null,
      DisplayName: 'test2',
      ExpiryDate: null,
      IsBillingProfileRequired: true,
      OwningAppId: '/Guid(50785fde-3082-47ac-a36d-06282ac5c7da)/',
      OwningTenantId: '/Guid(00000000-0000-0000-0000-000000000000)/',
      Region: null,
      ResourceGroup: null,
      SPContainerTypeBillingClassification: 0
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    auth.connection.spoUrl = spoAdminUrl.replace('-admin.sharepoint.com', '.sharepoint.com');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId, name: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither id nor name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if name is passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerTypeName });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the container type', async () => {
    await command.action(logger, { options: { id: containerTypeId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the container type when prompt is not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves([]);

    await command.action(logger, { options: { name: containerTypeName } });
    assert(postStub.notCalled);
  });

  it('correctly removes a container type by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            SchemaVersion: '15.0.0.0',
            LibraryVersion: '16.0.25919.12007',
            ErrorInfo: null,
            TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
          },
          7,
          {
            IsNull: false
          }
        ];
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerTypeId, force: true, verbose: true } });
    assert.strictEqual(postStub.firstCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
  });

  it('correctly removes a container type by name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (postStub.callCount === 1) {
          return [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 46, {
              "IsNull": false
            }, 47, CsomContainerTypeResponse
          ];
        }
        if (postStub.callCount === 2) {
          return [
            {
              SchemaVersion: '15.0.0.0',
              LibraryVersion: '16.0.25919.12007',
              ErrorInfo: null,
              TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
            },
            7,
            {
              IsNull: false
            }
          ];
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerTypeName, verbose: true, force: true } });
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
  });

  it('correctly throws error when retrieving container types', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
              "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
            }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
          }
        ];
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { name: containerTypeName, force: true } }), new CommandError('An error has occurred'));
  });

  it('correctly throws error when container type was not found by name', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (postStub.callCount === 1) {
          return [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 46, {
              "IsNull": false
            }, 47, CsomContainerTypeResponse
          ];
        }
        if (postStub.callCount === 2) {
          return [
            {
              SchemaVersion: '15.0.0.0',
              LibraryVersion: '16.0.25919.12007',
              ErrorInfo: null,
              TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
            },
            7,
            {
              IsNull: false
            }
          ];
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { name: 'nonexistent', force: true } }),
      new CommandError(`The specified container type 'nonexistent' does not exist.`));
  });

  it('correctly removes a container type by name when multiple containers with the same name exist', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        if (postStub.callCount === 1) {
          return [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
            }, 46, {
              "IsNull": false
            }, 47, [
              ...CsomContainerTypeResponse,
              {
                _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
                ApplicationRedirectUrl: 'https://contoso.sharepoint.com/redirect',
                AzureSubscriptionId: '/Guid(11111111-2222-3333-4444-555555555555)/',
                ContainerTypeId: '/Guid(aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee)/',
                CreationDate: '2025-01-15T10:30:00Z',
                DisplayName: containerTypeName,
                ExpiryDate: '2026-01-15T10:30:00Z',
                IsBillingProfileRequired: false,
                OwningAppId: '/Guid(12345678-90ab-cdef-1234-567890abcdef)/',
                OwningTenantId: '/Guid(99999999-8888-7777-6666-555555555555)/',
                Region: 'Europe',
                ResourceGroup: 'ContosoResourceGroup',
                SPContainerTypeBillingClassification: 2
              }
            ]
          ];
        }
        if (postStub.callCount === 2) {
          return [
            {
              SchemaVersion: '15.0.0.0',
              LibraryVersion: '16.0.25919.12007',
              ErrorInfo: null,
              TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
            },
            7,
            {
              IsNull: false
            }
          ];
        }
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(CsomContainerTypeResponse.find(c => c.DisplayName === containerTypeName));

    await command.action(logger, { options: { name: containerTypeName, verbose: true, force: true } });
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
  });

  it('correctly handles error when removing a container type', async () => {
    const errorMessage = `Tenant 7d858e1d-a366-48d1-8a15-ce45a733916b cannot delete Container Type ${containerTypeId} as it is a DirectToCustomer Container Type.`;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            SchemaVersion: '15.0.0.0',
            LibraryVersion: '16.0.25919.12007',
            ErrorInfo: {
              ErrorMessage: errorMessage,
              ErrorValue: null,
              TraceCorrelationId: 'cd4a91a1-6041-c000-29c0-26f4566b5b74',
              ErrorCode: -2146232832,
              ErrorTypeName: 'Microsoft.SharePoint.SPException'
            },
            TraceCorrelationId: 'cd4a91a1-6041-c000-29c0-26f4566b5b74'
          }
        ];
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { id: containerTypeId, force: true } }),
      new CommandError(errorMessage));
  });
});