import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../Auth';
import { Logger } from '../cli/Logger';
import request from '../request';
import { sinonUtil } from '../utils/sinonUtil';
import { FormDigestInfo, spo } from '../utils/spo';

const stubPostResponses: any = (
  folderAddResp: any = null
) => {
  return sinon.stub(request, 'post').callsFake((opts) => {
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath') > -1) {
      if (folderAddResp) {
        return folderAddResp;
      }
      else {
        return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "4t4", "ProgID": null, "ServerRelativeUrl": "/sites/JohnDoe/Shared Documents/4t4", "TimeCreated": "2018-10-26T22:50:27Z", "TimeLastModified": "2018-10-26T22:50:27Z", "UniqueId": "3f5428e2-b0a8-4d35-87df-89621ed5b457", "WelcomePage": "" });
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
        return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "f", "ProgID": null, "ServerRelativeUrl": "/sites/JohnDoe/Shared Documents/4t4/f", "TimeCreated": "2018-10-26T22:54:19Z", "TimeLastModified": "2018-10-26T22:54:20Z", "UniqueId": "0d680f20-53da-4516-b3f6-ed98b1d928e8", "WelcomePage": "" });
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

  it('retrieves tenant app catalog url', async () => {
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP_TenantSettings_Current') {
        return Promise.resolve({ CorporateCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' });
      }

      return Promise.reject('Invalid request');
    });


    const tenantAppCatalogUrl = await spo.getTenantAppCatalogUrl(logger, false);
    assert.deepEqual(tenantAppCatalogUrl, 'https://contoso.sharepoint.com/sites/appcatalog');
  });

  it('returns null when tenant app catalog not configured', async () => {
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP_TenantSettings_Current') {
        return Promise.resolve({ CorporateCatalogUrl: null });
      }

      return Promise.reject('Invalid request');
    });

    const tenantAppCatalogUrl = await spo.getTenantAppCatalogUrl(logger, false);
    assert.deepEqual(tenantAppCatalogUrl, null);
  });

  it('handles error when retrieving SPO URL failed while retrieving tenant app catalog url', (done) => {
    const errorMessage = 'Couldn\'t retrieve SharePoint URL';
    auth.service.spoUrl = undefined;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP_TenantSettings_Current') > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject(errorMessage);
    });

    spo
      .getTenantAppCatalogUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, errorMessage);
          done();
        }
        catch (e) {
          done(e);
        }
      });
  });

  it('handles error when retrieving the tenant app catalog URL fails', (done) => {
    const errorMessage = 'Couldn\'t retrieve tenant app catalog URL';
    auth.service.spoUrl = 'https://contoso.sharepoint.com';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP_TenantSettings_Current') > -1) {
        return Promise.reject(errorMessage);
      }

      return Promise.reject('Invalid request');
    });

    spo
      .getTenantAppCatalogUrl(logger, false)
      .then(() => {
        done('Expected error');
      }, (err: string) => {
        try {
          assert.strictEqual(err, errorMessage);
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

  //# region Custom Action Mock Responses
  const customActionOnSiteResponse1 = { "ClientSideComponentId": "d1e5e0d6-109d-40c4-a53e-924073fe9bbd", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "a6c7bef2-42d5-405c-a89f-6e36b3c302b3", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 2, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnSiteResponse2 = { "ClientSideComponentId": "230edcf5-2df5-480f-9707-ae1118726912", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "06d3eebb-6e30-4346-aecd-f84a342a9316", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 2, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnWebResponse1 = { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnWebResponse2 = { "ClientSideComponentId": "a405a600-7a21-49e7-9964-5e8b010b9eec", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };

  const customActionsOnSiteResponse = [customActionOnSiteResponse1, customActionOnSiteResponse2];
  const customActionsOnWebResponse = [customActionOnWebResponse1, customActionOnWebResponse2];
  //# endregion

  it(`returns a list of custom actions with scope 'All'`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: customActionsOnSiteResponse });
      }
      else if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: customActionsOnWebResponse });
      }

      return Promise.reject('Invalid request');
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'All');
    assert.deepEqual(customActions, [
      ...customActionsOnSiteResponse,
      ...customActionsOnWebResponse
    ]);
  });

  it(`returns a list of custom actions with scope 'Site'`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: customActionsOnSiteResponse });
      }

      return Promise.reject('Invalid request');
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Site');
    assert.deepEqual(customActions, customActionsOnSiteResponse);
  });

  it(`returns a list of custom actions with scope 'Web'`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: customActionsOnWebResponse });
      }

      return Promise.reject('Invalid request');
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Web');
    assert.deepEqual(customActions, customActionsOnWebResponse);
  });

  it(`returns a list of custom actions with scope 'Web' with a filter`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Web/UserCustomActions?$filter=ClientSideComponentId eq guid'b41916e7-e69d-467f-b37f-ff8ecf8f99f2'`) > -1) {
        return Promise.resolve({ value: [customActionOnWebResponse1] });
      }

      return Promise.reject('Invalid request');
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Web', `ClientSideComponentId eq guid'b41916e7-e69d-467f-b37f-ff8ecf8f99f2'`);
    assert.deepEqual(customActions, [customActionOnWebResponse1]);
  });

  it(`retrieves a custom action by id with scope 'All'`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Site/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return Promise.resolve(customActionOnSiteResponse1);
      }
      else if ((opts.url as string).indexOf(`/_api/Web/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return Promise.resolve({ 'odata.null': true });
      }

      return Promise.reject('Invalid request');
    });

    const customAction = await spo.getCustomActionById('https://contoso.sharepoint.com/sites/sales', 'd1e5e0d6-109d-40c4-a53e-924073fe9bbd');
    assert.deepEqual(customAction, customActionOnSiteResponse1);
  });

  it(`retrieves a custom action by id with scope 'Site'`, async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/Site/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return Promise.resolve(customActionOnSiteResponse1);
      }

      return Promise.reject('Invalid request');
    });

    const customAction = await spo.getCustomActionById('https://contoso.sharepoint.com/sites/sales', 'd1e5e0d6-109d-40c4-a53e-924073fe9bbd', 'Site');
    assert.deepEqual(customAction, customActionOnSiteResponse1);
  });

  it(`retrieves Azure AD ID by SPO user ID sucessfully`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/siteusers/GetById('9')?$select=AadObjectId`) {
        return {
          AadObjectId: {
            NameId: '6cc1797e-5463-45ec-bb1a-b93ec198bab6',
            NameIdIssuer: 'urn:federation:microsoftonline'
          }
        };
      }

      throw 'Invalid request';
    });

    const customAction = await spo.getUserAzureIdBySpoId('https://contoso.sharepoint.com/sites/sales', '9');
    assert.deepEqual(customAction, '6cc1797e-5463-45ec-bb1a-b93ec198bab6');
  });

  it(`throws error retrieving a custom action by id with a wrong scope value`, async () => {
    try {
      await spo.getCustomActionById('https://contoso.sharepoint.com/sites/sales', 'd1e5e0d6-109d-40c4-a53e-924073fe9bbd', 'Invalid');
      assert.fail('Expected an error to be thrown');
    }
    catch (e) {
      assert.deepEqual(e, `Invalid scope 'Invalid'. Allowed values are 'Site', 'Web' or 'All'.`);
    }
  });

  it(`throws error retrieving a list of custom actions with a wrong scope value`, async () => {
    try {
      await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Invalid');
      assert.fail('Expected an error to be thrown');
    }
    catch (e) {
      assert.deepEqual(e, `Invalid scope 'Invalid'. Allowed values are 'Site', 'Web' or 'All'.`);
    }
  });

  //#region Navigation menu state responses
  const topNavigationResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2039', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2041', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Sub level 1', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2040', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/PnPCoreSDKTestGroup', 'Title': 'Site B', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2001', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': '/sites/team-a/sitepages/about.aspx', 'Title': 'About', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1025', 'StartingNodeTitle': 'Quick launch', 'Version': '2023-03-09T18:33:53.5468097Z' };
  const quickLaunchResponse = { 'AudienceIds': [], 'FriendlyUrlPrefix': '', 'IsAudienceTargetEnabledForGlobalNav': false, 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2003', 'Nodes': [{ 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2006', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': '/sites/SharePointDemoSite#/', 'Title': 'Sub Item', 'Translations': [] }], 'NodeType': 0, 'OpenInNewWindow': true, 'SimpleUrl': 'http://google.be', 'Title': 'Site A', 'Translations': [] }, { 'AudienceIds': [], 'CurrentLCID': 1033, 'CustomProperties': [], 'FriendlyUrlSegment': '', 'IsDeleted': false, 'IsHidden': false, 'IsTitleForExistingLanguage': false, 'Key': '2018', 'Nodes': [], 'NodeType': 0, 'OpenInNewWindow': null, 'SimpleUrl': 'https://google.be', 'Title': 'Site B', 'Translations': [] }], 'SimpleUrl': '', 'SPSitePrefix': '/sites/SharePointDemoSite', 'SPWebPrefix': '/sites/SharePointDemoSite', 'StartingNodeKey': '1002', 'StartingNodeTitle': 'SharePoint Top Navigation Bar', 'Version': '2023-03-09T18:34:53.650545Z' };
  const webUrl = 'https://contoso.sharepoint.com/sites/sales';
  //#endregion

  it(`retrieves the quick launch navigation response`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return quickLaunchResponse;
      }

      throw 'Invalid request';
    });

    const quickLaunch = await spo.getQuickLaunchMenuState(webUrl);
    assert.deepEqual(quickLaunch, quickLaunchResponse);
  });

  it(`retrieves the top navigation response`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return topNavigationResponse;
      }

      throw 'Invalid request';
    });

    const topNavigation = await spo.getTopNavigationMenuState(webUrl);
    assert.deepEqual(topNavigation, topNavigationResponse);
  });

  it(`saves the menu state for the top navigation`, async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/navigation/MenuState`) {
        return topNavigationResponse;
      }

      if (opts.url === `${webUrl}/_api/navigation/SaveMenuState`) {
        return;
      }

      throw 'Invalid request';
    });

    const topNavigation = await spo.getTopNavigationMenuState(webUrl);
    await spo.saveMenuState(webUrl, topNavigation);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { menuState: topNavigation });
  });
});