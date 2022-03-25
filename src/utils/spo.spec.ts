import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../Auth';
import { Logger } from '../cli';
import request from '../request';
import { FormDigestInfo, sinonUtil, spo } from '../utils';

const stubPostResponses: any = (
  folderAddResp: any = null
) => {
  return sinon.stub(request, 'post').callsFake((opts) => {
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath') > -1) {
      if (folderAddResp) {
        return folderAddResp;
      }
      else {
        return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "4t4", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/4t4", "TimeCreated": "2018-10-26T22:50:27Z", "TimeLastModified": "2018-10-26T22:50:27Z", "UniqueId": "3f5428e2-b0a8-4d35-87df-89621ed5b457", "WelcomePage": "" });
      }

    }
    return Promise.reject('Invalid request');
  });
};

const stubGetResponses: any = (
  getFolderByServerRelativeUrlResp: any = null
) => {
  return sinon.stub(request, 'get').callsFake((opts) => {
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
      if (getFolderByServerRelativeUrlResp) {
        return getFolderByServerRelativeUrlResp;
      }
      else {
        return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "f", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/4t4/f", "TimeCreated": "2018-10-26T22:54:19Z", "TimeLastModified": "2018-10-26T22:54:20Z", "UniqueId": "0d680f20-53da-4516-b3f6-ed98b1d928e8", "WelcomePage": "" });
      }
    }
    return Promise.reject('Invalid request');
  });
};

describe('utils/spo', () => {
  let logger: Logger;
  let log: string[];
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo,
      spo.getSpoAdminUrl,
      spo.getRequestDigest
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('reuses current digestcontext when expireat is a future date', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    spo
      .ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false)
      .then((formDigest) => {
        try {
          assert.notStrictEqual(typeof formDigest, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, err => done(err));
  });

  it('reuses current digestcontext when expireat is a future date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    spo
      .ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true)
      .then((formDigest) => {
        try {
          assert.notStrictEqual(typeof formDigest, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, err => done(err));
  });

  it('retrieves new digestcontext when no context present', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    spo
      .ensureFormDigest('https://contoso.sharepoint.com', logger, undefined, false)
      .then(ctx => {
        try {
          assert.notStrictEqual(typeof ctx, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, e => {
        done(e);
      });
  });

  it('retrieves updated digestcontext when expireat is past date', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc',
          FormDigestTimeoutSeconds: 1800,
          FormDigestExpiresAt: new Date(),
          WebFullUrl: 'https://contoso.sharepoint.com'
        });
      }
      return Promise.reject('Invalid request');
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    spo
      .ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false)
      .then(ctx => {
        try {
          assert.notStrictEqual(typeof ctx, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, err => done(err));
  });

  it('retrieves updated digestcontext when expireat is past date (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }
      return Promise.reject('Invalid request');
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    spo
      .ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false)
      .then(ctx => {
        try {
          assert.notStrictEqual(typeof ctx, 'undefined');
          done();
        }
        catch (e) {
          done(e);
        }
      }, err => done(err));
  });

  it('handles error when contextinfo could not be retrieved (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.reject('Invalid request');
      }
      return Promise.reject('Invalid request');
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true).catch((err?: any) => {
      try {
        assert(err === "Invalid request");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    spo
      .getSpoUrl(logger, false)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously (debug)', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    spo
      .getSpoUrl(logger, true)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns retrieved SPO URL when persisting connection info failed', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject());
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.resolve({ webUrl: 'https://contoso.sharepoint.com' });
      }

      return Promise.reject('Invalid request');
    });

    spo
      .getSpoUrl(logger, false)
      .then((spoUrl: string) => {
        try {
          assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving SPO URL failed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    spo
      .getSpoUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving SPO admin URL failed', (done) => {
    sinon.stub(spo, 'getSpoUrl').callsFake(() => Promise.reject('An error has occurred'));

    spo
      .getSpoAdminUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves tenant ID when not retrieved previously', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    }));

    spo
      .getTenantId(logger, false)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('retrieves tenant ID when not retrieved previously (debug)', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    }));

    spo
      .getTenantId(logger, true)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns retrieved tenant ID when persisting connection info failed', (done) => {
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return Promise.resolve(JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]));
      }

      return Promise.reject('Invalid request');
    });
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    }));

    spo
      .getTenantId(logger, false)
      .then((tenantId: string) => {
        try {
          assert.strictEqual(tenantId, 'tenantId');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('returns error when retrieving tenant ID failed', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    sinon.stub(spo, 'getSpoAdminUrl').callsFake(() => Promise.resolve('https://contoso-admin.sharepoint.com'));
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    }));

    spo
      .getTenantId(logger, false)
      .then(() => {
        done('Error expected');
      }, (err: any) => {
        try {
          assert.strictEqual(err, 'An error has occurred');
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('should reject if wrong url param', (done) => {
    spo
      .ensureFolder("abc", "abc", logger, true)
      .then(() => {
        done('Should reject, not resolve');
      }, (err: any) => {
        assert.strictEqual(err, 'webFullUrl is not a valid URL');
        done();
      });
  });

  it('should reject if empty folder param', (done) => {
    spo
      .ensureFolder("https://contoso.sharepoint.com", "", logger, true)
      .then(() => {
        done('Should reject, not resolve');
      }, (err: any) => {
        assert.strictEqual(err, 'folderToEnsure cannot be empty');
        done();
      });
  });

  it('should handle folder creation failure', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });

    const expectedError = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Cannot create folder." } } });

    const folderCreationErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses(folderCreationErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, false)
      .then(() => {
        done('Should not resolve, but reject');
      }, (err: any) => {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(expectedError));
        done();
      });
  });

  it('should handle folder creation failure (debug)', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });

    const expectedError = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Cannot create folder." } } });

    const folderCreationErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses(folderCreationErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, true)
      .then(() => {
        done('Should not resolve, but reject');
      }, (err: any) => {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(expectedError));
        done();
      });
  });

  it('should succeed in adding folder if it does not exist (debug)', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses();

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, true)
      .then(() => {
        assert.strictEqual(loggerLogSpy.lastCall.args[0], 'All sub-folders exist');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should succeed in adding folder if it does not exist', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses();

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, false)
      .then(() => {
        assert.strictEqual(loggerLogSpy.notCalled, true);
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should succeed if all folders exist (debug)', (done) => {
    stubPostResponses();
    stubGetResponses();

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, true)
      .then(() => {
        assert.strictEqual(loggerLogSpy.called, true);
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should succeed if all folders exist', (done) => {
    stubPostResponses();
    stubGetResponses();

    spo
      .ensureFolder("https://contoso.sharepoint.com", "abc", logger, false)
      .then(() => {
        assert.strictEqual(loggerLogSpy.called, false);
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should have the correct url when calling AddSubFolderUsingPath (POST)', (done) => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com", "/folder2/folder3", logger, true)
      .then(() => {
        assert.strictEqual(postStubs.lastCall.args[0].url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Ffolder2%27&@a2=%27folder3%27');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should have the correct url including uppercase letters when calling AddSubFolderUsingPath', (done) => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true)
      .then(() => {
        assert.strictEqual(postStubs.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should call two times AddSubFolderUsingPath when folderUrl is folder2/folder3', (done) => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true)
      .then(() => {
        assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%27&@a2=%27folder2%27');
        assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should handle end slashes in the command options for webUrl and for folder', (done) => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com/sites/Site1/", "/folder2/folder3/", logger, true)
      .then(() => {
        assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%27&@a2=%27folder2%27');
        assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should have the correct url when folder option has uppercase letters when calling AddSubFolderUsingPath', (done) => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    });
    stubGetResponses(folderDoesNotExistErrorResp);

    spo
      .ensureFolder("https://contoso.sharepoint.com/sites/site1/", "PnP1/Folder2/", logger, true)
      .then(() => {
        assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2Fsite1%27&@a2=%27PnP1%27');
        assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2Fsite1%2FPnP1%27&@a2=%27Folder2%27');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  it('should call GetFolderByServerRelativeUrl with the correct url OData values', (done) => {
    stubPostResponses();
    const getStubs: sinon.SinonStub = stubGetResponses();

    spo
      .ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true)
      .then(() => {
        assert.strictEqual(getStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2FSite1%2Ffolder2\')');
        assert.strictEqual(getStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2FSite1%2Ffolder2%2Ffolder3\')');
        done();
      }, (err: any) => {
        done(err);
      });
  });
});