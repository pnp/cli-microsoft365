import * as sinon from 'sinon';
import * as assert from 'assert';
import _request from './request';
import * as requestPromise from 'request-promise-native';
import Utils from './Utils';
import * as https from 'https';
import request = require('request');
import auth from './Auth';
import { ClientRequest } from 'http';

describe('Request', () => {
  const cmdInstance = {
    commandWrapper: {
      command: 'command'
    },
    log: (msg: any) => { },
    prompt: () => { },
    action: () => { }
  };

  let _options: requestPromise.OptionsWithUrl;

  beforeEach(() => {
    _request.cmd = cmdInstance;
    _request.debug = false;
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => Promise.resolve('ABC'));
  });

  afterEach(() => {
    _request.debug = false;
    Utils.restore([
      global.setTimeout,
      https.request,
      (_request as any).req,
      cmdInstance.log,
      auth.ensureAccessToken
    ]);
  });

  it('fails when no command instance set', (done) => {
    _request.cmd = undefined as any;
    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        done('Error expected');
      }, (err: any) => {
        try {
          assert.strictEqual(err, 'Command reference not set on the request object');
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it('sets user agent on all requests', (done) => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        done('Error expected');
      }, () => {
        try {
          assert((_options.headers as request.Headers)['user-agent'].indexOf('NONISV|SharePointPnP|Office365CLI') > -1);
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it('uses gzip compression on all requests', (done) => {
    https.request
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        done('Error expected');
      }, () => {
        try {
          assert((_options.headers as request.Headers)['accept-encoding'].indexOf('gzip') > -1);
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it('sets access token on all requests', (done) => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/',
        headers: {}
      })
      .then(() => {
        done('Error expected');
      }, () => {
        try {
          assert((_options.headers as request.Headers)['authorization'].indexOf('Bearer ABC') > -1);
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it(`doesn't set access token on anonymous requests`, (done) => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/',
        headers: {
          'x-anonymous': true
        }
      })
      .then(() => {
        done('Error expected');
      }, () => {
        try {
          assert.strictEqual(typeof (_options.headers as request.Headers)['authorization'], 'undefined');
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it(`removes the anonymous header on anonymous requests`, (done) => {
    sinon.stub(https, 'request').callsFake((options: any) => {
      _options = options;
      return new ClientRequest('', () => { });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/',
        headers: {
          'x-anonymous': true
        }
      })
      .then(() => {
        done('Error expected');
      }, () => {
        try {
          assert.strictEqual(typeof (_options.headers as request.Headers)['x-anonymous'], 'undefined');
          done();
        }
        catch (err) {
          done(err);
        }
      });
  });

  it('sets method to GET for a GET request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'GET');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('sets method to HEAD for a HEAD request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .head({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'HEAD');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('sets method to POST for a POST request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .post({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'POST');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('sets method to PATCH for a PATCH request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .patch({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'PATCH');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('sets method to PUT for a PUT request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .put({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'PUT');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('sets method to DELETE for a DELETE request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .delete({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(_options.method, 'DELETE');
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err) => {
        done(err);
      });
  });

  it('returns response of a successful GET request', (done) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.resolve();
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        done();
      }, (err) => {
        done(err);
      });
  });

  it('correctly handles failed GET request', (cb) => {
    sinon.stub(_request as any, 'req').callsFake((options) => {
      _options = options;
      return Promise.reject('Error');
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        cb('Error expected');
      }, (err) => {
        try {
          assert.strictEqual(err, 'Error');
          cb();
        }
        catch (e) {
          cb(e);
        }
      });
  });

  it('repeats 429-throttled request after the designated retry value', (done) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {
              'retry-after': 60
            }
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(timeout, 60000);
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err) => {
        done(err);
      });
  });

  it('repeats 429-throttled request after 10s if no value specified', (done) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(timeout, 10000);
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err) => {
        done(err);
      });
  });

  it('repeats 429-throttled request after 10s if the specified value is not a number', (done) => {
    let i: number = 0;
    let timeout: number = -1;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {
              'retry-after': 'a'
            }
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      timeout = to;
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(timeout, 10000);
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err) => {
        done(err);
      });
  });

  it('repeats 429-throttled request until it succeeds', (done) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(i, 4);
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err: any) => {
        done(err);
      });
  });

  it('repeats 503-throttled request until it succeeds', (done) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ < 3) {
        return Promise.reject({
          response: {
            statusCode: 503,
            headers: {}
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert.strictEqual(i, 4);
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err: any) => {
        done(err);
      });
  });

  it('correctly handles request that was first 429-throttled and then failed', (done) => {
    let i: number = 0;

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {}
          }
        })
      }
      else {
        return Promise.reject('Error');
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        done('Expected error')
      }, (err) => {
        try {
          assert.strictEqual(err, 'Error');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('logs additional info for throttled requests in debug mode', (done) => {
    let i: number = 0;
    _request.debug = true;
    const logSpy: sinon.SinonSpy = sinon.spy(cmdInstance, 'log');

    sinon.stub(_request as any, 'req').callsFake(() => {
      if (i++ === 0) {
        return Promise.reject({
          response: {
            statusCode: 429,
            headers: {
              'retry-after': 10
            }
          }
        })
      }
      else {
        return Promise.resolve();
      }
    });
    sinon.stub(global as NodeJS.Global, 'setTimeout').callsFake((fn, to) => {
      fn();
      return {} as any;
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert(logSpy.calledWith('Request throttled. Waiting 10sec before retrying...'));
          done();
        }
        catch (err) {
          done(err);
        }
      }, (err: any) => {
        done(err);
      });
  });

  it('logs response body in debug mode', (done) => {
    _request.debug = true;
    const logSpy: sinon.SinonSpy = sinon.spy(cmdInstance, 'log');

    sinon.stub(_request as any, 'req').callsFake(() => {
      return Promise.resolve({
        hello: 'world'
      });
    });

    _request
      .get({
        url: 'https://contoso.sharepoint.com/'
      })
      .then(() => {
        try {
          assert(logSpy.calledWith(JSON.stringify({
            hello: 'world'
          })));
          done();
        }
        catch (err) {
          done(err)
        }
      }, (err: any) => {
        done(err);
      });
  });
})