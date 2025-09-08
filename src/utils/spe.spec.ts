
import assert from 'assert';
import sinon from 'sinon';
import { spe } from './spe.js';
import { sinonUtil } from './sinonUtil.js';
import request from '../request.js';
import auth from '../Auth.js';
import { cli } from '../cli/cli.js';
import { odata } from './odata.js';
import { formatting } from './formatting.js';

describe('utils/spe', () => {
  const containerTypeResponse = [
    {
      id: 'de988700-d700-020e-0a00-0831f3042f00',
      name: 'Container Type 1',
      owningAppId: '11335700-9a00-4c00-84dd-0c210f203f00',
      billingClassification: 'trial',
      createdDateTime: '01/20/2025',
      expirationDateTime: '02/20/2025',
      etag: 'RVRhZw==',
      settings: {
        urlTemplate: 'https://app.contoso.com/redirect?tenant={tenant-id}&drive={drive-id}&folder={folder-id}&item={item-id}',
        isDiscoverabilityEnabled: true,
        isSearchEnabled: true,
        isItemVersioningEnabled: true,
        itemMajorVersionLimit: 50,
        maxStoragePerContainerInBytes: 104857600,
        isSharingRestricted: false,
        consumingTenantOverridables: ''
      }
    },
    {
      id: '88aeae-d700-020e-0a00-0831f3042f01',
      name: 'Container Type 2',
      owningAppId: '33225700-9a00-4c00-84dd-0c210f203f01',
      billingClassification: 'standard',
      createdDateTime: '01/20/2025',
      expirationDateTime: null,
      etag: 'RVRhZw==',
      settings: {
        urlTemplate: '',
        isDiscoverabilityEnabled: true,
        isSearchEnabled: true,
        isItemVersioningEnabled: false,
        itemMajorVersionLimit: 100,
        maxStoragePerContainerInBytes: 104857600,
        isSharingRestricted: false,
        consumingTenantOverridables: ''
      }
    }
  ];

  before(() => {
    auth.connection.active = true;
  });

  afterEach(() => {
    sinonUtil.restore([
      odata.getAllItems,
      request.get,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
    sinon.restore();
  });

  it('correctly retrieves the container type ID by name when using getContainerTypeIdByName', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$select=id,name&$filter=name eq '${formatting.encodeQueryParameter('Container Type 1')}'`) {
        return [
          containerTypeResponse[0]
        ];
      }

      throw 'Invalid GET request ' + url;
    });

    const actual = await spe.getContainerTypeIdByName('Container Type 1');
    assert.strictEqual(actual, 'de988700-d700-020e-0a00-0831f3042f00');
  });

  it('correctly throws error when name not found when using getContainerTypeIdByName', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$select=id,name&$filter=name eq '${formatting.encodeQueryParameter('Container Type 5')}'`) {
        return [];
      }

      throw 'Invalid GET request ' + url;
    });

    await assert.rejects(spe.getContainerTypeIdByName('Container Type 5'),
      new Error(`The specified container type 'Container Type 5' does not exist.`));
  });

  it('correctly handles multiple results when using getContainerTypeIdByName', async () => {
    sinon.stub(odata, 'getAllItems').callsFake(async (url) => {
      if (url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$select=id,name&$filter=name eq '${formatting.encodeQueryParameter('Container Type 1')}'`) {
        return containerTypeResponse;
      }

      throw 'Invalid GET request ' + url;
    });

    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(containerTypeResponse[0]);
    const actual = await spe.getContainerTypeIdByName('Container Type 1');

    assert(stubMultiResults.calledOnce);
    assert.strictEqual(actual, 'de988700-d700-020e-0a00-0831f3042f00');
  });

  it('correctly gets a container by its name using getContainerIdByName', async () => {
    const containerTypeId = '0e95d161-d90d-4e3f-8b94-788a6b40aa48';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: [
            {
              id: 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
              displayName: 'My File Storage Container'
            },
            {
              id: 'b!t18F8ybsHUq1z3LTz8xvZqP8zaSWjkFNhsME-Fepo75dTf9vQKfeRblBZjoSQrd7',
              displayName: 'My File Storage Container 2'
            }
          ]
        };
      }

      throw 'Invalid GET request:' + opts.url;
    });

    const actual = await spe.getContainerIdByName(containerTypeId, 'my FILE storage Container 2');
    assert.strictEqual(actual, 'b!t18F8ybsHUq1z3LTz8xvZqP8zaSWjkFNhsME-Fepo75dTf9vQKfeRblBZjoSQrd7');
  });

  it('correctly throws error when container was not found using getContainerIdByName', async () => {
    const containerTypeId = '0e95d161-d90d-4e3f-8b94-788a6b40aa48';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: [
            {
              id: 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
              displayName: 'My File Storage Container'
            },
            {
              id: 'b!t18F8ybsHUq1z3LTz8xvZqP8zaSWjkFNhsME-Fepo75dTf9vQKfeRblBZjoSQrd7',
              displayName: 'My File Storage Container 2'
            }
          ]
        };
      }

      throw 'Invalid GET request:' + opts.url;
    });

    await assert.rejects(spe.getContainerIdByName(containerTypeId, 'nonexistent container'),
      new Error(`The specified container 'nonexistent container' does not exist.`));
  });

  it('correctly handles multiple results when using getContainerIdByName', async () => {
    const containerTypeId = '0e95d161-d90d-4e3f-8b94-788a6b40aa48';
    const containers = [
      {
        id: 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
        displayName: 'My File Storage Container'
      },
      {
        id: 'b!t18F8ybsHUq1z3LTz8xvZqP8zaSWjkFNhsME-Fepo75dTf9vQKfeRblBZjoSQrd7',
        displayName: 'My File Storage Container 2'
      },
      {
        id: 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND',
        displayName: 'My File Storage Container'
      }
    ];

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}&$select=id,displayName`) {
        return {
          value: containers
        };
      }

      throw 'Invalid GET request:' + opts.url;
    });

    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(containers.find(c => c.id === 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND')!);

    const actual = await spe.getContainerIdByName(containerTypeId, 'My File Storage Container');
    assert(stubMultiResults.calledOnce);
    assert.strictEqual(actual, 'b!McTeU0-dW0GxKwECWdW04TIvEK-Js9xJib_RFqF-CqZxNe3OHVAIT4SqBxGm4fND');
  });
});