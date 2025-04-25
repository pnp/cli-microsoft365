import assert from 'assert';
import { AxiosError } from 'axios';
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

  it('returns response on error', async () => {
    const url = 'https://example.com/api/resource';
    const options = {
      method: 'POST',
      headers: {
        'some-header': 'some-value'
      },
      body: JSON.stringify({ key: 'value' })
    };
    sinon.stub(request, 'execute').callsFake(async () => {
      throw new AxiosError('Request failed', 'ERR_BAD_REQUEST', undefined, undefined, {
        status: 400,
        statusText: 'Bad Request',
        headers: {
          'Content-Type': 'application/json',
          "expires": -1
        },
        config: {} as any,
        data: JSON.stringify({
          "error": "authorization_pending",
          "error_description": "AADSTS70016: OAuth 2.0 device flow error. Authorization is pending. Continue polling. Trace ID: 5ae9c106-cbfd-4d15-9cc1-dad6de7e2d00 Correlation ID: e95d8a70-e1c2-4693-a2f0-5441d290fbd0 Timestamp: 2025-04-25 09:05:06Z",
          "error_codes": [
            70016
          ],
          "timestamp": "2025-04-25 09:05:06Z",
          "trace_id": "5ae9c106-cbfd-4d15-9cc1-dad6de7e2d00",
          "correlation_id": "e95d8a70-e1c2-4693-a2f0-5441d290fbd0",
          "error_uri": "https://login.microsoftonline.com/error?code=70016"
        })
      });
    });
    const response = await msalNetworkClient.sendPostRequestAsync(url, options);
    assert.strictEqual(response.status, 400, 'Status code should be 400');
  });

  it('returns response on error without response information', async () => {
    const url = 'https://example.com/api/resource';
    const options = {
      method: 'POST',
      headers: {
        'some-header': 'some-value'
      },
      body: JSON.stringify({ key: 'value' })
    };
    sinon.stub(request, 'execute').callsFake(async () => {
      throw new AxiosError('Request failed', 'ERR_BAD_REQUEST', undefined, undefined, undefined);
    });
    const response = await msalNetworkClient.sendPostRequestAsync(url, options);
    assert.strictEqual(response.status, 400, 'Status code should be 400');
  });
});