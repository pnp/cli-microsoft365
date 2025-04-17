
import assert from 'assert';
import sinon from 'sinon';
import { spe } from './spe.js';
import { sinonUtil } from './sinonUtil.js';
import request from '../request.js';
import auth from '../Auth.js';
import config from '../config.js';
import { cli } from '../cli/cli.js';

describe('utils/spe', () => {
  const siteUrl = 'https://contoso.sharepoint.com';
  const adminUrl = siteUrl.replace('.sharepoint.com', '-admin.sharepoint.com');

  const containerTypeResponse = [
    {
      _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
      ApplicationRedirectUrl: null,
      AzureSubscriptionId: '/Guid(00000000-0000-0000-0000-000000000000)/',
      ContainerTypeId: '/Guid(073269af-f1d2-042d-2ef5-5bdd6ac83115)/',
      CreationDate: null,
      DisplayName: 'test1',
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
    auth.connection.active = true;
    auth.connection.spoUrl = siteUrl;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
    sinon.restore();
  });

  it('correctly retrieves a list of container types', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
          }, 46, {
            "IsNull": false
          }, 47, containerTypeResponse
        ];
      }

      throw 'Invalid request';
    });

    await spe.getAllContainerTypes(adminUrl);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="46" ObjectPathId="45" /><Method Name="GetSPOContainerTypes" Id="47" ObjectPathId="45"><Parameters><Parameter Type="Enum">1</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="45" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
  });

  it('correctly outputs a list of container types', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
          }, 46, {
            "IsNull": false
          }, 47, containerTypeResponse
        ];
      }

      throw 'Invalid request';
    });

    const actual = await spe.getAllContainerTypes(adminUrl);
    assert.deepStrictEqual(actual, [
      {
        ApplicationRedirectUrl: null,
        AzureSubscriptionId: '00000000-0000-0000-0000-000000000000',
        ContainerTypeId: '073269af-f1d2-042d-2ef5-5bdd6ac83115',
        CreationDate: null,
        DisplayName: 'test1',
        ExpiryDate: null,
        IsBillingProfileRequired: true,
        OwningAppId: 'df4085cc-9a38-4255-badc-5c5225610475',
        OwningTenantId: '00000000-0000-0000-0000-000000000000',
        Region: null,
        ResourceGroup: null,
        SPContainerTypeBillingClassification: 0
      },
      {
        ApplicationRedirectUrl: null,
        AzureSubscriptionId: '00000000-0000-0000-0000-000000000000',
        ContainerTypeId: '880ab3bd-5b68-01d4-3744-01a7656cf2ba',
        CreationDate: null,
        DisplayName: 'test2',
        ExpiryDate: null,
        IsBillingProfileRequired: true,
        OwningAppId: '50785fde-3082-47ac-a36d-06282ac5c7da',
        OwningTenantId: '00000000-0000-0000-0000-000000000000',
        Region: null,
        ResourceGroup: null,
        SPContainerTypeBillingClassification: 0
      }
    ]);
  });

  it('correctly throws error when retrieving container types', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
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

    await assert.rejects(spe.getAllContainerTypes(adminUrl), new Error('An error has occurred'));
  });

  it('correctly retrieves the container type ID by name when using getContainerTypeIdByName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
          }, 46, {
            "IsNull": false
          }, 47, containerTypeResponse
        ];
      }

      throw 'Invalid request';
    });

    const actual = await spe.getContainerTypeIdByName(adminUrl, 'test2');
    assert.strictEqual(actual, '880ab3bd-5b68-01d4-3744-01a7656cf2ba');
  });

  it('correctly throws error when name not found when using getContainerTypeIdByName', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
          }, 46, {
            "IsNull": false
          }, 47, containerTypeResponse
        ];
      }

      throw 'Invalid request';
    });

    await assert.rejects(spe.getContainerTypeIdByName(adminUrl, 'nonexistent'),
      new Error(`The specified container type 'nonexistent' does not exist.`));
  });

  it('correctly handles multiple results when using getContainerTypeIdByName', async () => {
    const containerTypes = [
      ...containerTypeResponse,
      {
        _ObjectType_: 'Microsoft.Online.SharePoint.TenantAdministration.SPContainerTypeProperties',
        ApplicationRedirectUrl: null,
        AzureSubscriptionId: '/Guid(00000000-0000-0000-0000-000000000000)/',
        ContainerTypeId: '/Guid(4c8bc473-2d5a-474d-b2f3-fc60b7d39726)/',
        CreationDate: null,
        DisplayName: 'test1',
        ExpiryDate: null,
        IsBillingProfileRequired: true,
        OwningAppId: '/Guid(48cc3066-7f0d-4cb9-80fb-f7891069c0f9)/',
        OwningTenantId: '/Guid(00000000-0000-0000-0000-000000000000)/',
        Region: null,
        ResourceGroup: null,
        SPContainerTypeBillingClassification: 0
      }
    ];

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12005", "ErrorInfo": null, "TraceCorrelationId": "2d63d39f-3016-0000-a532-30514e76ae73"
          }, 46, {
            "IsNull": false
          }, 47, containerTypes
        ];
      }

      throw 'Invalid request';
    });

    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(containerTypes.find(c => c.ContainerTypeId === '/Guid(4c8bc473-2d5a-474d-b2f3-fc60b7d39726)/')!);
    const actual = await spe.getContainerTypeIdByName(adminUrl, 'test1');
    assert(stubMultiResults.calledOnce);
    assert.strictEqual(actual, '4c8bc473-2d5a-474d-b2f3-fc60b7d39726');
  });
});