import assert from 'assert';
import sinon from 'sinon';
import auth from '../Auth.js';
import { Logger } from '../cli/Logger.js';
import config from '../config.js';
import { RoleDefinition } from '../m365/spo/commands/roledefinition/RoleDefinition.js';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { FormDigestInfo, spo } from '../utils/spo.js';
import { aadGroup } from './aadGroup.js';
import { formatting } from './formatting.js';

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
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath(DecodedUrl=') > -1) {
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      auth.storeConnectionInfo,
      spo.getSpoAdminUrl,
      spo.getRequestDigest,
      spo.ensureFormDigest,
      spo.siteExistsInTheRecycleBin,
      spo.getSpoUrl,
      spo.getTenantId,
      global.setTimeout
    ]);
    auth.service.spoUrl = undefined;
    auth.service.tenantId = undefined;
  });

  after(() => {
    sinon.restore();
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
        assert.strictEqual(getStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2Fsites%2FSite1%2Ffolder2\')');
        assert.strictEqual(getStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2Fsites%2FSite1%2Ffolder2%2Ffolder3\')');
        done();
      }, (err: any) => {
        done(err);
      });
  });

  //#region Custom Action Mock Responses
  const customActionOnSiteResponse1 = { "ClientSideComponentId": "d1e5e0d6-109d-40c4-a53e-924073fe9bbd", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "a6c7bef2-42d5-405c-a89f-6e36b3c302b3", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 2, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnSiteResponse2 = { "ClientSideComponentId": "230edcf5-2df5-480f-9707-ae1118726912", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "06d3eebb-6e30-4346-aecd-f84a342a9316", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 2, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnWebResponse1 = { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };
  const customActionOnWebResponse2 = { "ClientSideComponentId": "a405a600-7a21-49e7-9964-5e8b010b9eec", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" };

  const customActionsOnSiteResponse = [customActionOnSiteResponse1, customActionOnSiteResponse2];
  const customActionsOnWebResponse = [customActionOnWebResponse1, customActionOnWebResponse2];
  //#endregion

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

  it(`retrieves spo user by email sucessfully`, async () => {
    const userResponse = {
      Id: 11,
      IsHiddenInUI: false,
      LoginName: 'i:0#.f|membership|john.doe@contoso.com',
      Title: 'John Doe',
      PrincipalType: 1,
      Email: 'john.doe@contoso.com',
      Expiration: '',
      IsEmailAuthenticationGuestUser: false,
      IsShareByEmailGuestUser: false,
      IsSiteAdmin: false,
      UserId: {
        NameId: '10032002473c5ae3',
        NameIdIssuer: 'urn:federation:microsoftonline'
      },
      UserPrincipalName: 'john.doe@contoso.com'
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter('john.doe@contoso.com')}')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    const user = await spo.getUserByEmail('https://contoso.sharepoint.com/sites/sales', 'john.doe@contoso.com', logger, true);
    assert.deepEqual(user, userResponse);
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

  it(`retrieves spo group by name sucessfully`, async () => {
    const groupResponse = {
      Id: 11,
      IsHiddenInUI: false,
      LoginName: "groupname",
      Title: "groupname",
      PrincipalType: 8,
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      Description: "",
      OnlyAllowMembersViewMembership: true,
      OwnerTitle: "John Doe",
      RequestToJoinLeaveEmailSetting: null
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/sitegroups/GetByName('${formatting.encodeQueryParameter('groupname')}')`) {
        return groupResponse;
      }

      throw 'Invalid request';
    });

    const group = await spo.getGroupByName('https://contoso.sharepoint.com/sites/sales', 'groupname', logger, true);
    assert.deepEqual(group, groupResponse);
  });

  it(`retrieves roledefinition by name sucessfully`, async () => {
    const roledefinitionResponse: RoleDefinition = {
      BasePermissions: {
        High: 176,
        Low: 138612833
      },
      Description: "Can view pages and list items and download documents.",
      Hidden: false,
      Id: 1073741827,
      Name: "Read",
      Order: 128,
      RoleTypeKind: 2,
      BasePermissionsValue: [
        "ViewListItems",
        "OpenItems",
        "ViewVersions",
        "ViewFormPages",
        "Open",
        "ViewPages",
        "CreateSSCSite",
        "BrowseUserInfo",
        "UseClientIntegration",
        "UseRemoteAPIs",
        "CreateAlerts"
      ],
      RoleTypeKindValue: "Reader"
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/roledefinitions`) {
        return { value: [roledefinitionResponse] };
      }

      throw 'Invalid request';
    });

    const roledefintion = await spo.getRoleDefinitionByName('https://contoso.sharepoint.com/sites/sales', 'Read', logger, true);
    assert.deepEqual(roledefintion, roledefinitionResponse);
  });

  it(`handles error when no roledefinition by name is found`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/roledefinitions`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(spo.getRoleDefinitionByName('https://contoso.sharepoint.com/sites/sales', 'Read', logger, true), 'An error occured');
  });

  it('checks successfully if site exists', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "c340489e-80cc-5000-c5b4-01b2ce71e9bf"
            }, 197, {
              "IsNull": false
            }, 199, {
              "IsNull": false
            }, 200, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SiteProperties", "_ObjectIdentity_": "c340489e-80cc-5000-c5b4-01b2ce71e9bf|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "AllowDownloadingNonWebViewableFiles": true, "AllowEditing": true, "AllowSelfServiceUpgrade": true, "AverageResourceUsage": 0, "CommentsOnSitePagesDisabled": false, "CompatibilityLevel": 15, "ConditionalAccessPolicy": 0, "CurrentResourceUsage": 0, "DenyAddAndCustomizePages": 1, "DisableAppViews": 2, "DisableCompanyWideSharingLinks": 2, "DisableFlows": 2, "HasHolds": false, "LastContentModifiedDate": "\/Date(2018,1,7,19,9,58,513)\/", "Lcid": 1033, "LockIssue": null, "LockState": "Unlock", "NewUrl": "", "Owner": "admin@contoso.onmicrosoft.com", "OwnerEmail": "admin@contoso.onmicrosoft.com", "PWAEnabled": 0, "RestrictedToRegion": 3, "SandboxedCodeActivationCapability": 2, "SharingAllowedDomainList": "", "SharingBlockedDomainList": "", "SharingCapability": 0, "SharingDomainRestrictionMode": 0, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SiteDefinedSharingCapability": 0, "Status": "Active", "StorageMaximumLevel": 26214400, "StorageQuotaType": null, "StorageUsage": 2, "StorageWarningLevel": 25574400, "Template": "STS#0", "TimeZoneId": 4, "Title": "Team", "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0, "UserCodeWarningLevel": 0, "WebsCount": 1
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    const actual = await spo.siteExists('https://contoso.sharepoint.com', logger, true);
    assert.deepEqual(actual, false);
  });

  it('handles error when checking site', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    sinon.stub(spo, 'siteExistsInTheRecycleBin').resolves(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.siteExists('https://contoso.sharepoint.com', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'An error has occurred.');
    }
  });

  it('handles no site exception and checks for site in recycle bin', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    sinon.stub(spo, 'siteExistsInTheRecycleBin').resolves(true);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="197" ObjectPathId="196" /><ObjectPath Id="199" ObjectPathId="198" /><Query Id="200" ObjectPathId="198"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="196" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="198" ParentId="196" Name="GetSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Cannot get site https:\u002f\u002fcontoso.sharepoint.com", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoNoSiteException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    const actual = await spo.siteExists('https://contoso.sharepoint.com', logger, true);
    assert.deepEqual(actual, true);
  });

  it('checks succesfully if site exists in recycle bin', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-8041-5000-8242-77f6c560fa5e"
            }, 181, {
              "IsNull": false
            }, 182, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.DeletedSiteProperties", "_ObjectIdentity_": "e13c489e-8041-5000-8242-77f6c560fa5e|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nDeletedSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam", "DaysRemaining": 30, "DeletionTime": "\/Date(2018,1,7,18,57,20,530)\/", "SiteId": "\/Guid(cb09f194-0ee7-4c48-a44f-8c112fff4d4e)\/", "Status": "Recycled", "StorageMaximumLevel": 26214400, "Url": "https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam", "UserCodeMaximumLevel": 0
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    const actual = await spo.siteExistsInTheRecycleBin('https://contoso.sharepoint.com', logger, true);
    assert.deepEqual(actual, true);
  });

  it('handles no site in recycle bin exception because of unknown error', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    const actual = await spo.siteExistsInTheRecycleBin('https://contoso.sharepoint.com', logger, true);
    assert.deepEqual(actual, false);
  });

  it('handles error when checking site is in recycle bin', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="181" ObjectPathId="180" /><Query Id="182" ObjectPathId="180"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Method Id="180" ParentId="175" Name="GetDeletedSitePropertiesByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return (JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "e13c489e-2026-5000-8242-7ec96d02ba1d"
            }
          ]));
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.siteExistsInTheRecycleBin('https://contoso.sharepoint.com', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'An error has occurred.');
    }
  });

  it('deletes a site from the recycle bin succesfully', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "e13c489e-304e-5000-8242-705e26a87302"
            }, 185, {
              "IsNull": false
            }, 186, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": true, "PollingInterval": 15000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    await spo.deleteSiteFromTheRecycleBin('https://contoso.sharepoint.com', logger, true);
    assert(postStub.called);
  });

  it('handles an exception when trying to delete a site from the recycle bin', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.deleteSiteFromTheRecycleBin('https://contoso.sharepoint.com', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'An error has occurred.');
    }
  });

  it('deletes a site from the recycle bin succesfully and waits', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="185" ObjectPathId="184" /><Query Id="186" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Method Id="184" ParentId="175" Name="RemoveDeletedSite"><Parameters><Parameter Type="String">https://contoso.sharepoint.com</Parameter></Parameters></Method><Constructor Id="175" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]);
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await spo.deleteSiteFromTheRecycleBin('https://contoso.sharepoint.com', logger, true, true);
    assert(postStub.called);
  });

  it('adds a classic site with minimal options successfully', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">john.doe@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">undefined</Property><Property Name="Title" Type="String">team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 15000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    await spo.addSite('team', logger, true, false, 'ClassicSite', undefined, undefined, 'john.doe@contoso.com', undefined, false, undefined, undefined, undefined, 'https://contoso.sharepoint.com/sites/team', undefined, undefined, undefined, undefined, undefined, undefined);
    assert(postStub.called);
  });

  it('adds a classic site with full options successfully and waits for completing', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
    sinon.stub(spo, 'siteExists').resolves(true);
    sinon.stub(spo, 'deleteSiteFromTheRecycleBin').resolves();

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">john.doe@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">300</Property><Property Name="StorageWarningLevel" Type="Int64">275</Property><Property Name="Template" Type="String">PUBLISHING#0</Property><Property Name="TimeZoneId" Type="Int32">4</Property><Property Name="Title" Type="String">team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">100</Property><Property Name="UserCodeWarningLevel" Type="Double">90</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "d53a489e-c0c0-5000-58fc-d03b433dca89"
            }, 4, {
              "IsNull": false
            }, 6, {
              "IsNull": false
            }, 7, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Tenant", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nTenant", "AllowDownloadingNonWebViewableFiles": true, "AllowedDomainListForSyncClient": [

              ], "AllowEditing": true, "AllowLimitedAccessOnUnmanagedDevices": false, "BccExternalSharingInvitations": false, "BccExternalSharingInvitationsList": null, "BlockAccessOnUnmanagedDevices": false, "BlockDownloadOfAllFilesForGuests": false, "BlockDownloadOfAllFilesOnUnmanagedDevices": false, "BlockDownloadOfViewableFilesForGuests": false, "BlockDownloadOfViewableFilesOnUnmanagedDevices": false, "BlockMacSync": false, "CommentsOnSitePagesDisabled": false, "CompatibilityRange": "15,15", "ConditionalAccessPolicy": 0, "DefaultLinkPermission": 0, "DefaultSharingLinkType": 3, "DisableReportProblemDialog": false, "DisallowInfectedFileDownload": false, "DisplayNamesOfFileViewers": true, "DisplayStartASiteOption": true, "EmailAttestationReAuthDays": 30, "EmailAttestationRequired": false, "EnableGuestSignInAcceleration": false, "ExcludedFileExtensionsForSyncClient": [
                ""
              ], "ExternalServicesEnabled": true, "FileAnonymousLinkType": 2, "FilePickerExternalImageSearchEnabled": true, "FolderAnonymousLinkType": 2, "HideSyncButtonOnODB": false, "IPAddressAllowList": "", "IPAddressEnforcement": false, "IPAddressWACTokenLifetime": 15, "IsUnmanagedSyncClientForTenantRestricted": false, "IsUnmanagedSyncClientRestrictionFlightEnabled": true, "LegacyAuthProtocolsEnabled": true, "NoAccessRedirectUrl": null, "NotificationsInOneDriveForBusinessEnabled": true, "NotificationsInSharePointEnabled": true, "NotifyOwnersWhenInvitationsAccepted": true, "NotifyOwnersWhenItemsReshared": true, "ODBAccessRequests": 0, "ODBMembersCanShare": 0, "OfficeClientADALDisabled": false, "OneDriveForGuestsEnabled": false, "OneDriveStorageQuota": 1048576, "OptOutOfGrooveBlock": false, "OptOutOfGrooveSoftBlock": false, "OrphanedPersonalSitesRetentionPeriod": 30, "OwnerAnonymousNotification": true, "PermissiveBrowserFileHandlingOverride": false, "PreventExternalUsersFromResharing": false, "ProvisionSharedWithEveryoneFolder": false, "PublicCdnAllowedFileTypes": "CSS,EOT,GIF,ICO,JPEG,JPG,JS,MAP,PNG,SVG,TTF,WOFF", "PublicCdnEnabled": false, "PublicCdnOrigins": [

              ], "RequireAcceptingAccountMatchInvitedAccount": false, "RequireAnonymousLinksExpireInDays": 0, "ResourceQuota": 5300, "ResourceQuotaAllocated": 1200, "RootSiteUrl": "https:\u002f\u002fcontoso.sharepoint.com", "SearchResolveExactEmailOrUPN": false, "SharingAllowedDomainList": null, "SharingBlockedDomainList": null, "SharingCapability": 2, "SharingDomainRestrictionMode": 0, "ShowAllUsersClaim": true, "ShowEveryoneClaim": true, "ShowEveryoneExceptExternalUsersClaim": true, "ShowNGSCDialogForSyncOnODB": true, "ShowPeoplePickerSuggestionsForGuestUsers": false, "SignInAccelerationDomain": "", "SpecialCharactersStateInFileFolderNames": 1, "StartASiteFormUrl": null, "StorageQuota": 1061376, "StorageQuotaAllocated": 10669260800, "UseFindPeopleInPeoplePicker": false, "UsePersistentCookiesForExplorerView": false, "UserVoiceForFeedbackEnabled": true
            }, 8, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]);
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await spo.addSite('team', logger, true, true, 'ClassicSite', undefined, undefined, 'john.doe@contoso.com', undefined, true, undefined, undefined, 1033, 'https://contoso.sharepoint.com/sites/team', undefined, undefined, 4, 'PUBLISHING#0', 100, 90, 300, 275);
    assert(postStub.called);
  });

  it('handles exception when creating a classic site', async () => {
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="3"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="8" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="CreateSite"><Parameters><Parameter TypeId="{11f84fff-b8cf-47b6-8b50-34e692656606}"><Property Name="CompatibilityLevel" Type="Int32">0</Property><Property Name="Lcid" Type="UInt32">1033</Property><Property Name="Owner" Type="String">john.doe@contoso.com</Property><Property Name="StorageMaximumLevel" Type="Int64">100</Property><Property Name="StorageWarningLevel" Type="Int64">100</Property><Property Name="Template" Type="String">STS#0</Property><Property Name="TimeZoneId" Type="Int32">undefined</Property><Property Name="Title" Type="String">team</Property><Property Name="Url" Type="String">https://contoso.sharepoint.com/sites/team</Property><Property Name="UserCodeMaximumLevel" Type="Double">0</Property><Property Name="UserCodeWarningLevel" Type="Double">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "A site already exists at url https:\u002f\u002fcontoso.sharepoint.com\u002fsites\u002fteam.", "ErrorValue": null, "TraceCorrelationId": "c340489e-70f6-5000-c5b4-00bd039e3bf9", "ErrorCode": -2147024809, "ErrorTypeName": "System.ArgumentException"
              }, "TraceCorrelationId": "c340489e-70f6-5000-c5b4-00bd039e3bf9"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.addSite('team', logger, true, false, 'ClassicSite', undefined, undefined, 'john.doe@contoso.com', undefined, false, undefined, undefined, undefined, 'https://contoso.sharepoint.com/sites/team', undefined, undefined, undefined, undefined, undefined, undefined);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'A site already exists at url https://contoso.sharepoint.com/sites/team.');
    }
  });

  it('successfully creates a team site', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/GroupSiteManager/CreateGroupEx`) {
        return { SiteUrl: 'https://contoso.sharepoint.com/sites/team', ErrorMessage: null };
      }

      throw 'invalid request';
    });

    const actual = await spo.addSite('team', logger, true, false, 'TeamSite', 'team alias', 'team description', 'john.doe@contoso.com, sansa.stark@contoso.com', undefined, false, 'LBI', true, 1033);
    assert.deepStrictEqual(actual, 'https://contoso.sharepoint.com/sites/team');
  });

  it('handles exception when creating a team site', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/GroupSiteManager/CreateGroupEx`) {
        return { ErrorMessage: 'The teamsite already exists.' };
      }

      throw 'invalid request';
    });

    try {
      await spo.addSite('team', logger, true, false, 'TeamSite', 'team alias', undefined, 'john.doe@contoso.com, sansa.stark@contoso.com', undefined, false, undefined, true, 1033);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'The teamsite already exists.');
    }
  });

  it('successfully creates a communication site', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SPSiteManager/Create`) {
        return { SiteStatus: 2, SiteUrl: 'https://contoso.sharepoint.com/sites/team' };
      }

      throw 'invalid request';
    });

    const actual = await spo.addSite('team', logger, true, false, 'CommunicationSite', undefined, 'team description', 'john.doe@contoso.com', true, undefined, 'LBI', undefined, 1033, 'https://contoso.sharepoint.com/sites/team', undefined, '00000000-0000-0000-0000-000000000000');
    assert.deepStrictEqual(actual, 'https://contoso.sharepoint.com/sites/team');
  });

  it('handles exception when creating a communication site', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SPSiteManager/Create`) {
        return { SiteStatus: 0 };
      }

      throw 'invalid request';
    });

    try {
      await spo.addSite('team', logger, true, false, 'CommunicationSite', undefined, undefined, 'john.doe@contoso.com', true, undefined, undefined, undefined, 1033, 'https://contoso.sharepoint.com/sites/team');
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'An error has occurred while creating the site');
    }
  });

  it('successfully creates a communication site with site design Topic', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SPSiteManager/Create`) {
        return { SiteStatus: 2, SiteUrl: 'https://contoso.sharepoint.com/sites/team' };
      }

      throw 'invalid request';
    });

    const actual = await spo.addSite('team', logger, true, false, 'CommunicationSite', undefined, 'team description', 'john.doe@contoso.com', true, undefined, 'LBI', undefined, 1033, 'https://contoso.sharepoint.com/sites/team', 'Topic');
    assert.deepStrictEqual(actual, 'https://contoso.sharepoint.com/sites/team');
  });

  it('successfully creates a communication site with site design Showcase', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SPSiteManager/Create`) {
        return { SiteStatus: 2, SiteUrl: 'https://contoso.sharepoint.com/sites/team' };
      }

      throw 'invalid request';
    });

    const actual = await spo.addSite('team', logger, true, false, 'CommunicationSite', undefined, 'team description', 'john.doe@contoso.com', true, undefined, 'LBI', undefined, 1033, 'https://contoso.sharepoint.com/sites/team', 'Showcase');
    assert.deepStrictEqual(actual, 'https://contoso.sharepoint.com/sites/team');
  });

  it('successfully creates a communication site with site design Blank', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SPSiteManager/Create`) {
        return { SiteStatus: 2, SiteUrl: 'https://contoso.sharepoint.com/sites/team' };
      }

      throw 'invalid request';
    });

    const actual = await spo.addSite('team', logger, true, false, 'CommunicationSite', undefined, undefined, 'john.doe@contoso.com', true, undefined, undefined, undefined, 1033, 'https://contoso.sharepoint.com/sites/team', 'Blank');
    assert.deepStrictEqual(actual, 'https://contoso.sharepoint.com/sites/team');
  });

  it('applies a site design successfully', async () => {
    sinon.stub(spo, 'getSpoUrl').resolves('https://contoso.sharepoint.com');

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ApplySiteDesign`) {
        return { "ID": "4bfe70f8-f806-479c-9bf3-ffb2167b9ff5", "LogonName": "i:0#.f|membership|admin@contoso.onmicrosoft.com", "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e", "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575", "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf" };
      }

      throw 'invalid request';
    });

    await spo.applySiteDesign('9b142c22-037f-4a7f-9017-e9d8c0e34b98', 'https://contoso.sharepoint.com', logger, true);
    assert(postStub.called);
  });

  it('sets a site admin successfully', async () => {
    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">john.doe@contoso.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "b3d8499e-1079-5000-cb83-9da72405dfa6"
            }, 48, {
              "IsNull": false
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    await spo.setSiteAdmin('https://contoso.sharepoint.com', ctx, 'https://contoso.sharepoint.com/sites/team', 'john.doe@contoso.com', logger, true);
    assert(postStub.called);
  });

  it('handles exception when trying to update a site admin', async () => {
    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="48" ObjectPathId="47" /></Actions><ObjectPaths><Method Id="47" ParentId="34" Name="SetSiteAdmin"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/team</Parameter><Parameter Type="String">john.doe@contoso.com</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method><Constructor Id="34" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.setSiteAdmin('https://contoso.sharepoint.com', ctx, 'https://contoso.sharepoint.com/sites/team', 'john.doe@contoso.com', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'Unknown Error');
    }
  });

  it('sets groupified site admins successfully', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'john.doe@contoso.com' or userPrincipalName eq 'sansa.stark@contoso.com'&$select=id`) {
        return {
          value: [
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d728' },
            { id: 'b17ff355-cc97-4b90-9b46-e33d0d70d729' }
          ]
        };
      }

      throw 'invalid request';
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d728', principalName='')`) {
        return;
      }

      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SP.Directory.DirectorySession/Group('e10a459e-60c8-4000-8240-a68d6a12d39e')/Owners/Add(objectId='b17ff355-cc97-4b90-9b46-e33d0d70d729', principalName='')`) {
        return;
      }

      throw 'invalid request';
    });

    await spo.setGroupifiedSiteOwners('https://contoso-admin.sharepoint.com', 'e10a459e-60c8-4000-8240-a68d6a12d39e', 'john.doe@contoso.com,sansa.stark@contoso.com', logger, true);
    assert(postStub.called);
  });

  it('handles when site owners not found', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'john.doe@contoso.com' or userPrincipalName eq 'sansa.stark@contoso.com'&$select=id`) {
        return {
          value: []
        };
      }

      throw 'invalid request';
    });

    await spo.setGroupifiedSiteOwners('https://contoso-admin.sharepoint.com', 'e10a459e-60c8-4000-8240-a68d6a12d39e', 'john.doe@contoso.com,sansa.stark@contoso.com', logger, true);
    assert(getStub.called);
  });

  it('sets a group connected site successfully', async () => {
    sinon.stub(spo, 'getTenantId').resolves('a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant');
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    sinon.stub(aadGroup, 'setGroup').resolves();
    sinon.stub(spo, 'setGroupifiedSiteOwners').resolves();
    sinon.stub(spo, 'applySiteDesign').resolves();

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site?$select=GroupId,Id`) {
        return {
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        };
      }

      throw 'invalid request';
    });


    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPOGroup/UpdateGroupPropertiesBySiteId`) {
        return;
      }

      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Classification"><Parameter Type="String">HBI</Parameter></SetProperty><SetProperty Id="28" ObjectPathId="5" Name="DisableFlows"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="29" ObjectPathId="5" Name="ShareByEmailEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><SetProperty Id="30" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">0</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7331.1205", "ErrorInfo": null, "TraceCorrelationId": "54d8499e-b001-5000-cb83-9445b3944fb9"
            }, 14, {
              "IsNull": false
            }, 15, {
              "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSiteProperties\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam"
            }, 16, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 15000
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636540580851601240&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await spo.updateSite('https://contoso.sharepoint.com', logger, true, 'team', 'HBI', true, true, 'john.doe@contoso.com,sansa.stark@contoso.com', true, 'eb2f31da-9461-4fbf-9ea1-9959b134b89e', 'Disabled');
    assert(postStub.called);
  });

  it('sets a group connected site successfully with minmal options', async () => {
    sinon.stub(spo, 'getTenantId').resolves('a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant');
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site?$select=GroupId,Id`) {
        return {
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: 'e10a459e-60c8-4000-8240-a68d6a12d39e'
        };
      }

      throw 'invalid request';
    });


    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="SharingCapability"><Parameter Type="Enum">0</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    await spo.updateSite('https://contoso.sharepoint.com', logger, true, 'team', undefined, undefined, undefined, undefined, undefined, undefined, 'Disabled');
    assert(postStub.called);
  });

  it('sets a non group connected site successfully', async () => {
    sinon.stub(spo, 'getTenantId').resolves('a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant');
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    sinon.stub(spo, 'setSiteAdmin').resolves();

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site?$select=GroupId,Id`) {
        return {
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        };
      }

      throw 'invalid request';
    });


    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">team</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 5000
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    await spo.updateSite('https://contoso.sharepoint.com', logger, true, 'team', undefined, undefined, undefined, 'john.doe@contoso.com,sansa.stark@contoso.com');
    assert(postStub.called);
  });

  it('handles exception when using isPublic with a group connected site', async () => {
    sinon.stub(spo, 'getTenantId').resolves('a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant');
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site?$select=GroupId,Id`) {
        return {
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        };
      }

      throw 'invalid request';
    });

    try {
      await spo.updateSite('https://contoso.sharepoint.com', logger, true, 'team', undefined, undefined, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `The isPublic option can't be set on a site that is not groupified`);
    }
  });

  it('handles exception when updating site', async () => {
    sinon.stub(spo, 'getTenantId').resolves('a61d499e-50aa-5000-8242-7169ab88ce08|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;Tenant');
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/site?$select=GroupId,Id`) {
        return {
          Id: '255a50b2-527f-4413-8485-57f4c17a24d1',
          GroupId: '00000000-0000-0000-0000-000000000000'
        };
      }

      throw 'invalid request';
    });


    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><SetProperty Id="27" ObjectPathId="5" Name="Title"><Parameter Type="String">team</Parameter></SetProperty><ObjectPath Id="14" ObjectPathId="13" /><ObjectIdentityQuery Id="15" ObjectPathId="5" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="53d8499e-d0d2-5000-cb83-9ade5be42ca4|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SiteProperties&#xA;https%3A%2F%2Fcontoso.sharepoint.com" /><Method Id="13" ParentId="5" Name="Update" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
                "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
              }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
            }
          ]);
        }
      }

      throw 'invalid request';
    });

    try {
      await spo.updateSite('https://contoso.sharepoint.com', logger, true, 'team');
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'An error has occurred.');
    }
  });

  it(`retrieves web properties susccessfully`, async () => {
    const webResponse = {
      value: [{
        AllowRssFeeds: false,
        AlternateCssUrl: null,
        AppInstanceId: "00000000-0000-0000-0000-000000000000",
        Configuration: 0,
        Created: null,
        CurrentChangeToken: null,
        CustomMasterUrl: null,
        Description: null,
        DesignPackageId: null,
        DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
        EnableMinimalDownload: false,
        HorizontalQuickLaunch: false,
        Id: "d8d179c7-f459-4f90-b592-14b08e84accb",
        IsMultilingual: false,
        Language: 1033,
        LastItemModifiedDate: null,
        LastItemUserModifiedDate: null,
        MasterUrl: null,
        NoCrawl: false,
        OverwriteTranslationsOnChange: false,
        ResourcePath: null,
        QuickLaunchEnabled: false,
        RecycleBinEnabled: false,
        ServerRelativeUrl: null,
        SiteLogoUrl: null,
        SyndicationEnabled: false,
        Title: "Subsite",
        TreeViewEnabled: false,
        UIVersion: 15,
        UIVersionConfigurationEnabled: false,
        Url: "https://contoso.sharepoint.com/subsite",
        WebTemplate: "STS"
      }]
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web`) {
        return webResponse;
      }

      throw 'invalid request';
    });

    const actual = await spo.getWeb('https://contoso.sharepoint.com', logger, true);
    assert.deepStrictEqual(actual, webResponse);
  });


  it('removes a file', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('%2FSharedDocuments%2FDocument.docx')`) {
        return;
      }

      throw 'Invalid request';
    });


    await spo.removeFile('https://contoso.sharepoint.com', 'SharedDocuments/Document.docx');
    assert(postStub.called);
  });

  it('removes a file and recycles it', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativeUrl('%2FSharedDocuments%2FDocument.docx')/recycle()`) {
        return;
      }

      throw 'Invalid request';
    });


    await spo.removeFile('https://contoso.sharepoint.com', 'SharedDocuments/Document.docx', true, logger, true);
    assert(postStub.called);
  });
});