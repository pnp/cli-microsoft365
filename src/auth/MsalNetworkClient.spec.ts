import assert from 'assert';
import sinon from 'sinon';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { MsalNetworkClient } from './MsalNetworkClient.js';

describe('MsalNetworkClient', () => {
  const msalNetworkClient = new MsalNetworkClient();

  afterEach(() => {
    sinonUtil.restore([
      request.execute
    ]);
  });

  it('sends GET request', async () => {
    const url = 'https://example.com/api/resource';
    const options = {
      headers: {
        'some-header': 'some-value'
      }
    };
    sinon.stub(request, 'execute').callsFake(async opts => {
      if (opts.url === url &&
        opts.method === 'GET' &&
        opts.headers &&
        opts.headers['x-anonymous'] === true &&
        opts.headers['some-header'] === 'some-value') {
        return {
          status: 200,
          headers: {
            'Content-Type': 'application/json'
          },
          data: JSON.stringify({ success: true })
        };
      }

      throw `Invalid options: ${JSON.stringify(opts)}`;
    });
    const response = await msalNetworkClient.sendGetRequestAsync(url, options);
    assert.strictEqual(response.status, 200, 'Status code should be 200');
    assert.strictEqual((response.body as any).success, true, 'Response body should contain success: true');
    assert.strictEqual(response.headers['Content-Type'], 'application/json', 'Response should have Content-Type header');
  });

  it('sends POST request', async () => {
    const url = 'https://example.com/api/resource';
    const options = {
      headers: {
        'some-header': 'some-value'
      },
      body: JSON.stringify({ key: 'value' })
    };
    sinon.stub(request, 'execute').callsFake(async opts => {
      if (opts.url === url &&
        opts.method === 'POST' &&
        opts.headers &&
        opts.headers['x-anonymous'] === true &&
        opts.headers['some-header'] === 'some-value' &&
        opts.data === JSON.stringify({ key: 'value' })) {
        return {
          status: 200,
          headers: {
            'Content-Type': 'application/json',
            'x-another-header': 100
          },
          data: JSON.stringify({ success: true })
        };
      }

      throw `Invalid options: ${JSON.stringify(opts)}`;
    });
    const response = await msalNetworkClient.sendPostRequestAsync(url, options);
    assert.strictEqual(response.status, 200, 'Status code should be 200');
    assert.strictEqual((response.body as any).success, true, 'Response body should contain success: true');
    assert.strictEqual(response.headers['Content-Type'], 'application/json', 'Response should have Content-Type header');
  });
});