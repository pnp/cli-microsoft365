import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import auth from '../Auth';
import { powerPlatform } from './powerPlatform';
import { sinonUtil } from "./sinonUtil";

describe('utils/powerPlatform', () => {
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    auth.service.connected = true;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('returns correct dynamics url as admin', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    const actual = await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', true);
    assert.strictEqual(actual, 'https://contoso-dev.api.crm4.dynamics.com');
  });

  it('returns correct dynamics url', async () => {
    const envResponse: any = { "properties": { "linkedEnvironmentMetadata": { "instanceApiUrl": "https://contoso-dev.api.crm4.dynamics.com" } } };

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(envResponse);
        }
      }

      return Promise.reject('Invalid request');
    });

    const actual = await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', false);
    assert.strictEqual(actual, 'https://contoso-dev.api.crm4.dynamics.com');
  });

  it('handles no environment found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/environments/someRandomGuid?api-version=2020-10-01&$select=properties.linkedEnvironmentMetadata.instanceApiUrl`)) {
        throw Error('Random Error');
      }

      return 'Invalid request';
    });

    try {
      await powerPlatform.getDynamicsInstanceApiUrl('someRandomGuid', false);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The environment 'someRandomGuid' could not be retrieved. See the inner exception for more details: Random Error`));
    }
  });
});