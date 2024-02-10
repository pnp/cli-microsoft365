import assert from 'assert';
import { ClientRequest } from 'http';
import https from 'https';
import sinon from 'sinon';
import auth, { CloudType } from './Auth.js';
import { Logger } from './cli/Logger.js';
import _request, { CliRequestOptions } from './request.js';
import { sinonUtil } from './utils/sinonUtil.js';
import { timersUtil } from './utils/timersUtil.js';

describe('Request', () => {
  const logger: Logger = {
    log: async () => { },
    logRaw: async () => { },
    logToStderr: async () => { }
  };

  let _options: CliRequestOptions;
  const retryAfter = 10;

  beforeEach(() => {
    _request.logger = logger;
    _request.debug = false;
    sinon.stub(auth, 'ensureAccessToken').resolves('ABC');
  });

  afterEach(() => {
    _request.debug = false;
    sinonUtil.restore([
      process.env,
      https.request,
      (_request as any).req,
      logger.log,
      auth.ensureAccessToken,
      timersUtil.setTimeout
    ]);
  });

  it('fails when no command instance set', async () => {
    _request.logger = undefined as any;

    try {
      await _request.get({
        url: 'https://contoso.sharepoint.com/'
      });
      assert.fail('Error expected');
    }
    catch (err) {
      assert.strictEqual(err, 'Logger not set on the request object');
    }
  });

  it('sets user agent on all requests', async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/'
        });
      assert.fail('Error expected');
    }
    catch {
      assert((_options as any).headers['user-agent'].indexOf('NONISV|SharePointPnP|CLIMicrosoft365') > -1);
    }
  });

  it('uses gzip compression on all requests', async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/'
        });
      assert.fail('Error expected');
    }
    catch {
      assert((_options as any).headers['accept-encoding'].indexOf('gzip') > -1);
    }

  });

  it('sets access token on all requests', async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/',
          headers: {}
        });
      assert.fail('Error expected');
    }
    catch {
      assert((_options as any).headers['authorization'].indexOf('Bearer ABC') > -1);
    }
  });

  it(`doesn't set access token on anonymous requests`, async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request.get({
        url: 'https://contoso.sharepoint.com/',
        headers: {
          'x-anonymous': 'true'
        }
      });
      assert.fail('Error expected');
    }
    catch {
      assert.strictEqual(typeof (_options as any).headers['authorization'], 'undefined');
    }
  });

  it(`removes the anonymous header on anonymous requests`, async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/',
          headers: {
            'x-anonymous': 'true'
          }
        });
      assert.fail('Error expected');
    }
    catch {
      assert.strictEqual(typeof (_options as any).headers['x-anonymous'], 'undefined');
    }
  });


  it(`removes the resource header on distinguished resource requests`, async () => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/',
          headers: {
            'x-resource': 'https://contoso.sharepoint.com'
          }
        });
      assert.fail('Error expected');
    }
    catch {
      assert.strictEqual(typeof (_options as any).headers['x-resource'], 'undefined');
    }
  });

  it('sets method to GET for a GET request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
    assert.strictEqual(_options.method, 'GET');
  });

  it('sets method to HEAD for a HEAD request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .head({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(_options.method, 'HEAD');
  });

  it('sets method to POST for a POST request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .post({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(_options.method, 'POST');
  });

  it('sets method to PATCH for a PATCH request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .patch({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(_options.method, 'PATCH');
  });

  it('sets method to PUT for a PUT request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .put({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(_options.method, 'PUT');
  });

  it('sets method to DELETE for a DELETE request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .delete({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(_options.method, 'DELETE');
  });

  it('returns response of a successful GET request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
  });

  it('returns response of a successful GET request, with overridden authorization', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/',
        headers: {
          authorization: 'Bearer 123'
        }
      });
  });

  it('returns response of a successful GET request for large file (stream)', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      (options as CliRequestOptions).responseType = "stream";
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
  });

  it('returns response of a successful GET request, with a proxy url', async () => {
    let proxyConfigured = false;
    sinon.stub(process, 'env').value({ 'HTTPS_PROXY': 'http://proxy.contoso.com:8080' });

    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options as CliRequestOptions;
      proxyConfigured = !!_options.proxy &&
        _options.proxy.host === 'proxy.contoso.com' &&
        _options.proxy.port === 8080 &&
        _options.proxy.protocol === 'http';
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert(proxyConfigured);
  });

  it('returns response of a successful GET request, with a proxy url and defaults port to 80', async () => {
    let proxyConfigured = false;
    sinon.stub(process, 'env').value({ 'HTTPS_PROXY': 'http://proxy.contoso.com' });

    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options as CliRequestOptions;
      proxyConfigured = !!_options.proxy &&
        _options.proxy.host === 'proxy.contoso.com' &&
        _options.proxy.port === 80 &&
        _options.proxy.protocol === 'http';
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert(proxyConfigured);
  });

  it('returns response of a successful GET request, with a proxy url and defaults port to 443', async () => {
    let proxyConfigured = false;
    sinon.stub(process, 'env').value({ 'HTTPS_PROXY': 'https://proxy.contoso.com' });

    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options as CliRequestOptions;
      proxyConfigured = !!_options.proxy &&
        _options.proxy.host === 'proxy.contoso.com' &&
        _options.proxy.port === 443 &&
        _options.proxy.protocol === 'http';
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
    assert(proxyConfigured);
  });

  it('returns response of a successful GET request, with a proxy url with username and password', async () => {
    let proxyConfigured = false;
    sinon.stub(process, 'env').value({ 'HTTPS_PROXY': 'http://username:password@proxy.contoso.com:8080' });

    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options as CliRequestOptions;
      proxyConfigured = !!_options.proxy &&
        _options.proxy.host === 'proxy.contoso.com' &&
        _options.proxy.port === 8080 &&
        _options.proxy.protocol === 'http' &&
        _options.proxy.auth?.username === 'username' &&
        _options.proxy.auth?.password === 'password';
      return { data: {} };
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert(proxyConfigured);
  });

  it('correctly handles failed GET request', async () => {
    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      throw 'Error';
    });

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/'
        });
      assert.fail('Error expected');
    }
    catch (err) {
      assert.strictEqual(err, 'Error');
    }
  });

  it('repeats 429-throttled request after the designated retry value', async () => {
    let i: number = 0;
    let timeout: number = -1;
    const retryAfter = 60;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {
              'retry-after': retryAfter
            }
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').callsFake(async (value: any) => {
      timeout = value;
      return;
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(timeout, retryAfter * 1000);
  });

  it('repeats 429-throttled request after 10s if no value specified', async () => {
    let i: number = 0;
    let timeout: number | undefined = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {}
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').callsFake(async (value: any) => {
      timeout = value;
      return;
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(timeout, retryAfter * 1000);
  });

  it('repeats 429-throttled request after 10s if the specified value is not a number', async () => {
    let i: number = 0;
    let timeout: number | undefined = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {
              'retry-after': 'a'
            }
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').callsFake(async (value: any) => {
      timeout = value;
      return;
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(timeout, retryAfter * 1000);
  });

  it('repeats 429-throttled request until it succeeds', async () => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        throw {
          response: {
            status: 429,
            headers: {}
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').resolves();

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(i, 4);
  });

  it('repeats 429-throttled request after the designated retry value for large file (stream)', async () => {
    let i: number = 0;
    let timeout: number | undefined = -1;
    const retryAfter = 60;

    sinon.stub(_request as any, 'req').callsFake(options => {
      _options = options as CliRequestOptions;
      (options as CliRequestOptions).responseType = "stream";

      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {
              'retry-after': retryAfter
            }
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').callsFake(async (value: any) => {
      timeout = value;
      return;
    });

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
    assert.strictEqual(timeout, retryAfter * 1000);
  });

  it('repeats 503-throttled request until it succeeds', async () => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        throw {
          response: {
            status: 503,
            headers: {}
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').resolves();

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });

    assert.strictEqual(i, 4);
  });

  it('correctly handles request that was first 429-throttled and then failed', async () => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {}
          }
        };
      }
      else {
        throw 'Error';
      }
    });

    sinon.stub(timersUtil, 'setTimeout').resolves();

    try {
      await _request
        .get({
          url: 'https://contoso.sharepoint.com/'
        });
      assert.fail('Error expected');
    }
    catch (err) {
      assert.strictEqual(err, 'Error');
    }
  });

  it('logs additional info for throttled requests in debug mode', async () => {
    let i: number = 0;
    _request.debug = true;
    const logSpy: sinon.SinonSpy = sinon.spy(logger, 'log');

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        throw {
          response: {
            status: 429,
            headers: {
              'retry-after': 10
            }
          }
        };
      }
      else {
        return { data: {} };
      }
    });

    sinon.stub(timersUtil, 'setTimeout').resolves();

    await _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      });
    assert(logSpy.calledWith('Request throttled. Waiting 10 sec before retrying...'));
  });

  it(`updates the URL for the China cloud`, async () => {
    let url;
    auth.connection.cloudType = CloudType.China;
    sinon.stub(_request as any, 'req').callsFake((options: any) => {
      url = options.url;
      return { data: {} };
    });
    await _request.execute({
      url: 'https://graph.microsoft.com/v1.0/me'
    });
    assert.strictEqual(url, 'https://microsoftgraph.chinacloudapi.cn/v1.0/me');
  });

  it(`updates the URL for the USGov cloud`, async () => {
    let url;
    auth.connection.cloudType = CloudType.USGov;
    sinon.stub(_request as any, 'req').callsFake((options: any) => {
      url = options.url;
      return { data: {} };
    });
    await _request.execute({
      url: 'https://graph.microsoft.com/v1.0/me'
    });
    assert.strictEqual(url, 'https://graph.microsoft.com/v1.0/me');
  });

  it(`updates the URL for the USGovDoD cloud`, async () => {
    let url;
    auth.connection.cloudType = CloudType.USGovDoD;
    sinon.stub(_request as any, 'req').callsFake((options: any) => {
      url = options.url;
      return { data: {} };
    });
    await _request.execute({
      url: 'https://graph.microsoft.com/v1.0/me'
    });
    assert.strictEqual(url, 'https://dod-graph.microsoft.us/v1.0/me');
  });

  it(`updates the URL for the USGovHigh cloud`, async () => {
    let url;
    auth.connection.cloudType = CloudType.USGovHigh;
    sinon.stub(_request as any, 'req').callsFake((options: any) => {
      url = options.url;
      return { data: {} };
    });
    await _request.execute({
      url: 'https://graph.microsoft.com/v1.0/me'
    });
    assert.strictEqual(url, 'https://graph.microsoft.us/v1.0/me');
  });
});
