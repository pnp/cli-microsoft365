import os from 'os';
import assert from 'assert';
import sinon from 'sinon';
import auth from '../Auth.js';
import { Logger } from '../cli/Logger.js';
import config from '../config.js';
import { RoleDefinition } from '../m365/spo/commands/roledefinition/RoleDefinition.js';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { CreateFileCopyJobsNameConflictBehavior, FormDigestInfo, SpoOperation, spo, CreateFolderCopyJobsNameConflictBehavior } from '../utils/spo.js';
import { entraGroup } from './entraGroup.js';
import { formatting } from './formatting.js';
import { Group } from '@microsoft/microsoft-graph-types';
import { timersUtil } from './timersUtil.js';

const stubPostResponses: any = (
  folderAddResp: any = null
) => {
  return sinon.stub(request, 'post').callsFake(async (opts) => {
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath') > -1) {
      if (folderAddResp) {
        throw folderAddResp;
      }
      else {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "4t4", "ProgID": null, "ServerRelativeUrl": "/sites/JohnDoe/Shared Documents/4t4", "TimeCreated": "2018-10-26T22:50:27Z", "TimeLastModified": "2018-10-26T22:50:27Z", "UniqueId": "3f5428e2-b0a8-4d35-87df-89621ed5b457", "WelcomePage": "" };
      }

    }
    throw 'Invalid request';
  });
};

const stubGetResponses: any = (
  getFolderByServerRelativeUrlResp: any = null
) => {
  return sinon.stub(request, 'get').callsFake(async (opts) => {
    if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath(DecodedUrl=') > -1) {
      if (getFolderByServerRelativeUrlResp) {
        throw getFolderByServerRelativeUrlResp;
      }
      else {
        return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "f", "ProgID": null, "ServerRelativeUrl": "/sites/JohnDoe/Shared Documents/4t4/f", "TimeCreated": "2018-10-26T22:54:19Z", "TimeLastModified": "2018-10-26T22:54:20Z", "UniqueId": "0d680f20-53da-4516-b3f6-ed98b1d928e8", "WelcomePage": "" };
      }
    }
    throw 'Invalid request';
  });
};

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

const entraGroupResponse = {
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
  UserId: null,
  UserPrincipalName: null
};

const copyJobInfo = {
  EncryptionKey: "2by8+2oizihYOFqk02Tlokj8lWUShePAEE+WMuA9lzA=",
  JobId: "d812e5a0-d95a-4e4f-bcb7-d4415e88c8ee",
  JobQueueUri: "https://spoam1db1m020p4.queue.core.windows.net/2-1499-20240831-29533e6c72c6464780b756c71ea3fe92?sv=2018-03-28&sig=aX%2BNOkUimZ3f%2B%2BvdXI95%2FKJI1e5UE6TU703Dw3Eb5c8%3D&st=2024-08-09T00%3A00%3A00Z&se=2024-08-31T00%3A00%3A00Z&sp=rap",
  SourceListItemUniqueIds: [
    'c194762b-3f54-4f5f-9f5c-eba26084e29d'
  ]
};

describe('utils/spo', () => {
  let logger: Logger;
  let log: string[];
  let loggerLogSpy: sinon.SinonSpy;

  const webUrl = 'https://contoso.sharepoint.com/sites/sales';

  before(() => {
    auth.connection.active = true;
    sinon.stub(timersUtil, 'setTimeout').resolves();
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
      spo.getTenantId
    ]);
    auth.connection.spoUrl = undefined;
    auth.connection.spoTenantId = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('reuses current digestcontext when expireat is a future date', async () => {
    sinon.stub(request, 'post').rejects('Invalid request');

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    const formDigest = await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false);
    assert.notStrictEqual(typeof formDigest, 'undefined');
  });

  it('reuses current digestcontext when expireat is a future date (debug)', async () => {
    sinon.stub(request, 'post').rejects('Invalid request');

    const futureDate = new Date();
    futureDate.setSeconds(futureDate.getSeconds() + 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: futureDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    const formDigest = await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true);
    assert.notStrictEqual(typeof formDigest, 'undefined');
  });

  it('retrieves new digestcontext when no context present', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return { FormDigestValue: 'abc' };
      }
      throw 'Invalid request';
    });

    const ctx = await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, undefined, false);
    assert.notStrictEqual(typeof ctx, 'undefined');
  });

  it('retrieves updated digestcontext when expireat is past date', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return {
          FormDigestValue: 'abc',
          FormDigestTimeoutSeconds: 1800,
          FormDigestExpiresAt: new Date(),
          WebFullUrl: 'https://contoso.sharepoint.com'
        };
      }
      throw 'Invalid request';
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    const formCtx = await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false);
    assert.notStrictEqual(typeof formCtx, 'undefined');
  });

  it('retrieves updated digestcontext when expireat is past date (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return { FormDigestValue: 'abc' };
      }
      throw 'Invalid request';
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    const formCtx = await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, false);
    assert.notStrictEqual(typeof formCtx, 'undefined');
  });

  it('handles error when contextinfo could not be retrieved (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        throw 'Different error than Invalid request';
      }
      throw 'Invalid request';
    });

    const pastDate = new Date();
    pastDate.setSeconds(pastDate.getSeconds() - 1800);

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: pastDate,
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    try {
      await spo.ensureFormDigest('https://contoso.sharepoint.com', logger, ctx, true);
      assert.fail('No error message thrown');
    }
    catch (e) {
      assert.strictEqual(e, 'Different error than Invalid request');
    }
  });

  it('retrieves tenant app catalog url', async () => {
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP_TenantSettings_Current') {
        return { CorporateCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' };
      }

      throw 'Invalid request';
    });

    const tenantAppCatalogUrl = await spo.getTenantAppCatalogUrl(logger, false);
    assert.deepEqual(tenantAppCatalogUrl, 'https://contoso.sharepoint.com/sites/appcatalog');
  });

  it('returns null when tenant app catalog not configured', async () => {
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_api/SP_TenantSettings_Current') {
        return { CorporateCatalogUrl: null };
      }

      throw 'Invalid request';
    });

    const tenantAppCatalogUrl = await spo.getTenantAppCatalogUrl(logger, false);
    assert.deepEqual(tenantAppCatalogUrl, null);
  });

  it('handles error when retrieving SPO URL failed while retrieving tenant app catalog url', async () => {
    const errorMessage = 'Couldn\'t retrieve SharePoint URL';
    auth.connection.spoUrl = undefined;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP_TenantSettings_Current') > -1) {
        throw 'An error has occurred';
      }

      throw errorMessage;
    });

    try {
      await spo.getTenantAppCatalogUrl(logger, false);
      assert.fail('No error message thrown');
    }
    catch (e) {
      assert.strictEqual(e, errorMessage);
    }
  });

  it('handles error when retrieving the tenant app catalog URL fails', async () => {
    const errorMessage = 'Couldn\'t retrieve tenant app catalog URL';
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/SP_TenantSettings_Current') > -1) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });
    try {
      await spo.getTenantAppCatalogUrl(logger, false);
      assert.fail('No error message thrown');
    }
    catch (e) {
      assert.strictEqual(e, errorMessage);
    }
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously', async () => {
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com' };
      }

      throw 'Invalid request';
    });

    const spoUrl = await spo.getSpoUrl(logger, false);
    assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
  });

  it('retrieves SPO URL from MS Graph when not retrieved previously (debug)', async () => {
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com' };
      }

      throw 'Invalid request';
    });

    const spoUrl = await spo.getSpoUrl(logger, true);
    assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
  });

  it('returns retrieved SPO URL when persisting connection info failed', async () => {
    sinon.stub(auth, 'storeConnectionInfo').rejects();
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com' };
      }

      throw 'Invalid request';
    });

    const spoUrl = await spo.getSpoUrl(logger, true);
    assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
  });

  it('throws error when ', async () => {
    sinon.stub(auth, 'storeConnectionInfo').rejects();
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com' };
      }

      throw 'Invalid request';
    });

    const spoUrl = await spo.getSpoUrl(logger, true);
    assert.strictEqual(spoUrl, 'https://contoso.sharepoint.com');
  });

  it('returns error when retrieving SPO URL failed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl') {
        throw 'An error has occurred';
      }

      throw 'Invalid request';
    });

    try {
      await spo.getSpoUrl(logger, false);
      assert.fail('No error message thrown');
    }
    catch (err) {
      assert.strictEqual(err, 'An error has occurred');
    }
  });

  it('returns error when retrieving SPO admin URL failed', async () => {
    sinon.stub(spo, 'getSpoUrl').rejects(new Error('An error has occurred'));

    try {
      await spo.getSpoAdminUrl(logger, false);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(err.message, 'An error has occurred');
    }
  });

  it('retrieves tenant ID when not retrieved previously', async () => {
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]);
      }

      throw 'Invalid request';
    });
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    const tenantId = await spo.getTenantId(logger, false);
    assert.strictEqual(tenantId, 'tenantId');
  });

  it('retrieves tenant ID when not retrieved previously (debug)', async () => {
    sinon.stub(auth, 'storeConnectionInfo').resolves();
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]);
      }

      throw 'Invalid request';
    });
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    const tenantId = await spo.getTenantId(logger, true);
    assert.strictEqual(tenantId, 'tenantId');
  });

  it('returns retrieved tenant ID when persisting connection info failed', async () => {
    sinon.stub(auth, 'storeConnectionInfo').rejects('An error has occurred');
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{
          _ObjectIdentity_: 'tenantId'
        }]);
      }

      throw 'Invalid request';
    });
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    const tenantId = await spo.getTenantId(logger, true);
    assert.strictEqual(tenantId, 'tenantId');
  });

  it('returns error when retrieving tenant ID failed', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    try {
      await spo.getTenantId(logger, false);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(err.message, 'An error has occurred');
    }
  });

  it('should reject if wrong url param', async () => {
    try {
      await spo.ensureFolder("abc", "abc", logger, true);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(err.message, 'webFullUrl is not a valid URL');
    }
  });

  it('should reject if empty folder param', async () => {
    try {
      await spo.ensureFolder("https://contoso.sharepoint.com", "", logger, true);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(err.message, 'folderToEnsure cannot be empty');
    }
  });

  it('should handle folder creation failure', async () => {
    const expectedError = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Cannot create folder." } } });

    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    stubPostResponses(expectedError);

    try {
      await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, false);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(expectedError));
    }
  });

  it('should handle folder creation failure (debug)', async () => {
    const expectedError = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Cannot create folder." } } });

    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    stubPostResponses(expectedError);

    try {
      await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, true);
      assert.fail('No error message thrown');
    }
    catch (err: any) {
      assert.strictEqual(JSON.stringify(err), JSON.stringify(expectedError));
    }
  });

  it('should succeed in adding folder if it does not exist (debug)', async () => {
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    stubPostResponses();

    await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, true);
    assert.strictEqual(loggerLogSpy.lastCall.args[0], 'All sub-folders exist');
  });

  it('should succeed in adding folder if it does not exist', async () => {
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));
    stubPostResponses();

    await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, false);
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('should succeed if all folders exist (debug)', async () => {
    stubPostResponses();
    stubGetResponses();

    await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, true);
    assert.strictEqual(loggerLogSpy.called, true);
  });

  it('should succeed if all folders exist', async () => {
    stubPostResponses();
    stubGetResponses();

    await spo.ensureFolder("https://contoso.sharepoint.com", "abc", logger, false);
    assert.strictEqual(loggerLogSpy.called, false);
  });

  it('should have the correct url when calling AddSubFolderUsingPath (POST)', async () => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));

    await spo.ensureFolder("https://contoso.sharepoint.com", "/folder2/folder3", logger, true);
    assert.strictEqual(postStubs.lastCall.args[0].url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Ffolder2%27&@a2=%27folder3%27');
  });

  it('should have the correct url including uppercase letters when calling AddSubFolderUsingPath', async () => {
    const postStubs: sinon.SinonStub = stubPostResponses();

    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));

    await spo.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true);
    assert.strictEqual(postStubs.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
  });

  it('should call two times AddSubFolderUsingPath when folderUrl is folder2/folder3', async () => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));

    await spo.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true);
    assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%27&@a2=%27folder2%27');
    assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
  });

  it('should handle end slashes in the command options for webUrl and for folder', async () => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));

    await spo.ensureFolder("https://contoso.sharepoint.com/sites/Site1/", "/folder2/folder3/", logger, true);
    assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%27&@a2=%27folder2%27');
    assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2FSite1%2Ffolder2%27&@a2=%27folder3%27');
  });

  it('should have the correct url when folder option has uppercase letters when calling AddSubFolderUsingPath', async () => {
    const postStubs: sinon.SinonStub = stubPostResponses();
    stubGetResponses(JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not found." } } }));

    await spo.ensureFolder("https://contoso.sharepoint.com/sites/site1/", "PnP1/Folder2/", logger, true);
    assert.strictEqual(postStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2Fsite1%27&@a2=%27PnP1%27');
    assert.strictEqual(postStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1=%27%2Fsites%2Fsite1%2FPnP1%27&@a2=%27Folder2%27');
  });

  it('should call GetFolderByServerRelativeUrl with the correct url OData values', async () => {
    stubPostResponses();
    const getStubs: sinon.SinonStub = stubGetResponses();

    await spo.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3", logger, true);
    assert.strictEqual(getStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2Fsites%2FSite1%2Ffolder2\')');
    assert.strictEqual(getStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativePath(DecodedUrl=\'%2Fsites%2FSite1%2Ffolder2%2Ffolder3\')');
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return { value: customActionsOnSiteResponse };
      }
      else if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return { value: customActionsOnWebResponse };
      }

      throw 'Invalid request';
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'All');
    assert.deepEqual(customActions, [
      ...customActionsOnSiteResponse,
      ...customActionsOnWebResponse
    ]);
  });

  it(`returns a list of custom actions with scope 'Site'`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return { value: customActionsOnSiteResponse };
      }

      throw 'Invalid request';
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Site');
    assert.deepEqual(customActions, customActionsOnSiteResponse);
  });

  it(`returns a list of custom actions with scope 'Web'`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return { value: customActionsOnWebResponse };
      }

      throw 'Invalid request';
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Web');
    assert.deepEqual(customActions, customActionsOnWebResponse);
  });

  it(`returns a list of custom actions with scope 'Web' with a filter`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Web/UserCustomActions?$filter=ClientSideComponentId eq guid'b41916e7-e69d-467f-b37f-ff8ecf8f99f2'`) > -1) {
        return { value: [customActionOnWebResponse1] };
      }

      throw 'Invalid request';
    });

    const customActions = await spo.getCustomActions('https://contoso.sharepoint.com/sites/sales', 'Web', `ClientSideComponentId eq guid'b41916e7-e69d-467f-b37f-ff8ecf8f99f2'`);
    assert.deepEqual(customActions, [customActionOnWebResponse1]);
  });

  it(`retrieves a custom action by id with scope 'All'`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Site/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return customActionOnSiteResponse1;
      }
      else if ((opts.url as string).indexOf(`/_api/Web/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return { 'odata.null': true };
      }

      throw 'Invalid request';
    });

    const customAction = await spo.getCustomActionById('https://contoso.sharepoint.com/sites/sales', 'd1e5e0d6-109d-40c4-a53e-924073fe9bbd');
    assert.deepEqual(customAction, customActionOnSiteResponse1);
  });

  it(`retrieves a custom action by id with scope 'Site'`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_api/Site/UserCustomActions(guid'd1e5e0d6-109d-40c4-a53e-924073fe9bbd')`) > -1) {
        return customActionOnSiteResponse1;
      }

      throw 'Invalid request';
    });

    const customAction = await spo.getCustomActionById('https://contoso.sharepoint.com/sites/sales', 'd1e5e0d6-109d-40c4-a53e-924073fe9bbd', 'Site');
    assert.deepEqual(customAction, customActionOnSiteResponse1);
  });

  it(`retrieves Microsoft Entra ID by SPO user ID successfully`, async () => {
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

  it(`retrieves spo user by email successfully`, async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/siteusers/GetByEmail('${formatting.encodeQueryParameter('john.doe@contoso.com')}')`) {
        return userResponse;
      }

      throw 'Invalid request';
    });

    const user = await spo.getUserByEmail('https://contoso.sharepoint.com/sites/sales', 'john.doe@contoso.com', logger, true);
    assert.deepEqual(user, userResponse);
  });

  it('successfully returns a SharePoint user when calling ensureUser', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/web/EnsureUser') {
        return userResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const user = await spo.ensureUser('https://contoso.sharepoint.com/sites/sales', 'john.doe@contoso.com');
    assert.deepStrictEqual(user, userResponse);
  });

  it('successfully ensures a SharePoint user when calling ensureUser', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/web/EnsureUser') {
        return userResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.ensureUser('https://contoso.sharepoint.com/sites/sales', 'john.doe@contoso.com');
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: 'john.doe@contoso.com' });
  });

  it('successfully throws an error when calling ensureEntraGroup with a group that is not security enabled', async () => {
    const graphGroup: Group = {
      id: '38243edd-76c7-4d6d-9093-9e90e6e7e28a',
      displayName: 'Sales',
      securityEnabled: false,
      mailEnabled: false
    };

    await assert.rejects(spo.ensureEntraGroup('https://contoso.sharepoint.com/sites/sales', graphGroup),
      new Error('Cannot ensure a Microsoft Entra ID group that is not security enabled.'));
  });

  it('successfully outputs the ensured group when calling ensureEntraGroup', async () => {
    const graphGroup: Group = {
      id: '38243edd-76c7-4d6d-9093-9e90e6e7e28a',
      displayName: 'Sales',
      securityEnabled: true,
      mailEnabled: false
    };

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/web/EnsureUser') {
        return entraGroupResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    const group = await spo.ensureEntraGroup('https://contoso.sharepoint.com/sites/sales', graphGroup);
    assert.deepStrictEqual(group, entraGroupResponse);
  });

  it('successfully ensures security group when calling ensureEntraGroup', async () => {
    const graphGroup: Group = {
      id: '38243edd-76c7-4d6d-9093-9e90e6e7e28a',
      displayName: 'Sales',
      securityEnabled: true,
      mailEnabled: false
    };

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/web/EnsureUser') {
        return entraGroupResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.ensureEntraGroup('https://contoso.sharepoint.com/sites/sales', graphGroup);
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: `c:0t.c|tenant|${graphGroup.id}` });
  });

  it('successfully ensures M365 group when calling ensureEntraGroup', async () => {
    const graphGroup: Group = {
      id: '38243edd-76c7-4d6d-9093-9e90e6e7e28a',
      displayName: 'Sales',
      securityEnabled: true,
      mailEnabled: true
    };

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/web/EnsureUser') {
        return entraGroupResponse;
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.ensureEntraGroup('https://contoso.sharepoint.com/sites/sales', graphGroup);
    assert.deepStrictEqual(postStub.firstCall.args[0].data, { logonName: `c:0o.c|federateddirectoryclaimprovider|${graphGroup.id}` });
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "e13c489e-304e-5000-8242-705e26a87302|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nRemoveDeletedSite\n636536266495764941\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\ncb09f194-0ee7-4c48-a44f-8c112fff4d4e", "IsComplete": true, "PollingInterval": 0
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }

        // done
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'invalid request';
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536245073557362\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="d53a489e-c0c0-5000-58fc-d03b433dca89|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;CreateSite&#xA;636536245073557362&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nCreateSite\n636536251347192220\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'invalid request';
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
    sinon.stub(entraGroup, 'setGroup').resolves();
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": false, "PollingInterval": 0
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="188" ObjectPathId="184"><Query SelectAllProperties="false"><Properties><Property Name="IsComplete" ScalarProperty="true" /><Property Name="PollingInterval" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="184" Name="54d8499e-b001-5000-cb83-9445b3944fb9|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023&#xA;SpoOperation&#xA;SetSite&#xA;636540580851601240&#xA;https%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam&#xA;00000000-0000-0000-0000-000000000000" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": null, "TraceCorrelationId": "803b489e-9066-5000-58fc-dc40eb096913"
            }, 39, {
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
            }
          ]);
        }
      }

      throw 'invalid request';
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
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
              "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation", "_ObjectIdentity_": "803b489e-9066-5000-58fc-dc40eb096913|908bed80-a04a-4433-b4a0-883d9847d110:67753f63-bc14-4012-869e-f808a43fe023\nSpoOperation\nSetSite\n636540580851601240\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fteam\n00000000-0000-0000-0000-000000000000", "IsComplete": true, "PollingInterval": 0
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

  it(`retrieves web properties successfully`, async () => {
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

  it(`applies a retention label to list items successfully`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(spo.applyRetentionLabelToListItems('https://contoso.sharepoint.com/sites/project-x', 'Some label', 'https://contoso.sharepoint.com/sites/project-x/list', [1], logger, true));
  });

  it(`removes a retention label from list items successfully`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetComplianceTagOnBulkItems`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","itemIds":[1]}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(spo.removeRetentionLabelFromListItems('https://contoso.sharepoint.com/sites/project-x', 'https://contoso.sharepoint.com/sites/project-x/list', [1], logger, true));
  });

  it(`applies a default retention label to a list successfully`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"Some label","blockDelete":false,"blockEdit":false,"syncToItems":true}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(spo.applyDefaultRetentionLabelToList('https://contoso.sharepoint.com/sites/project-x', 'Some label', 'https://contoso.sharepoint.com/sites/project-x/list', true, logger, true));
  });

  it(`removes a default retention label from a list successfully`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/SP_CompliancePolicy_SPPolicyStoreProxy_SetListComplianceTag`
        && JSON.stringify(opts.data) === '{"listUrl":"https://contoso.sharepoint.com/sites/project-x/list","complianceTagValue":"","blockDelete":false,"blockEdit":false,"syncToItems":false}') {
        return;
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(spo.removeDefaultRetentionLabelFromList('https://contoso.sharepoint.com/sites/project-x', 'https://contoso.sharepoint.com/sites/project-x/list', logger, true));
  });

  it('returns the correct site ID for a valid site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      const expectedUrl = 'https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/?$select=id';
      if (opts.url === expectedUrl) {
        return { id: 'contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42' };
      }

      throw 'Invalid request';
    });

    const id = await spo.getSiteId('https://contoso.sharepoint.com', logger);

    assert.strictEqual(id, 'contoso.sharepoint.com,ea49a393-e3e6-4760-a1b2-e96539e15372,66e2861c-96d9-4418-a75c-0ed1bca68b42');
  });

  it('returns the folder server relative URL by URL', async () => {
    const serverRelativeUrl = '/sites/sales/shared documents/folder1';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='${formatting.encodeQueryParameter(serverRelativeUrl)}')?$select=ServerRelativeUrl`
      ) {
        return { ServerRelativeUrl: serverRelativeUrl };
      }

      throw 'Invalid request';
    });

    const url = await spo.getFolderServerRelativeUrl(webUrl, serverRelativeUrl, undefined, logger, true);

    assert.strictEqual(url, serverRelativeUrl);
  });

  it('returns the folder server relative URL by id', async () => {
    const serverRelativeUrl = '/sites/sales/shared documents/folder1';
    const folderId = 'f09c4efe-b8c0-4e89-a166-03418661b89b';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFolderById('${folderId}')?$select=ServerRelativeUrl`) {
        return { ServerRelativeUrl: serverRelativeUrl };
      }

      throw 'Invalid request';
    });

    const url = await spo.getFolderServerRelativeUrl(webUrl, undefined, folderId, logger, true);

    assert.strictEqual(url, serverRelativeUrl);
  });

  it(`get the file properties with the server relative url`, async () => {
    const fileResponse = {
      ListItemAllFields: {
        FileSystemObjectType: 0,
        Id: 4,
        ServerRedirectedEmbedUri: 'https://contoso.sharepoint.com/sites/sales/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
        ServerRedirectedEmbedUrl: 'https://contoso.sharepoint.com/sites/sales/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
        ContentTypeId: '0x0101008E462E3ACE8DB844B3BEBF9473311889',
        ComplianceAssetId: null,
        Title: null,
        ID: 4,
        Created: '2018-02-05T09:42:36',
        AuthorId: 1,
        Modified: '2018-02-05T09:44:03',
        EditorId: 1,
        OData__CopySource: null,
        CheckoutUserId: null,
        OData__UIVersionString: '3.0',
        GUID: '2054f49e-0f76-46d4-ac55-50e1c057941c'
      },
      CheckInComment: '',
      CheckOutType: 2,
      ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
      CustomizedPageStatus: 0,
      ETag: '\'{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\'',
      Exists: true,
      IrmEnabled: false,
      Length: 331673,
      Level: 1,
      LinkingUri: 'https://contoso.sharepoint.com/sites/sales/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b',
      LinkingUrl: 'https://contoso.sharepoint.com/sites/sales/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b',
      MajorVersion: 3,
      MinorVersion: 0,
      Name: 'Test1.docx',
      ServerRelativeUrl: '/sites/sales/Documents/Test1.docx',
      TimeCreated: '2018-02-05T08:42:36Z',
      TimeLastModified: '2018-02-05T08:44:03Z',
      Title: '',
      UIVersion: 1536,
      UIVersionLabel: '3.0',
      UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${webUrl}/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?$expand=ListItemAllFields&@f='%2Fsites%2Fsales%2FDocuments%2FTest1.docx'`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    const actual = await spo.getFileAsListItemByUrl(webUrl, '/sites/sales/Documents/Test1.docx', logger, true);
    assert.strictEqual(actual, fileResponse.ListItemAllFields);
  });

  it(`sets the list item with system update`, async () => {
    const listItemResponse = {
      Attachments: false,
      AuthorId: 3,
      ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
      Created: '2018-03-15T10:43:10Z',
      EditorId: 3,
      GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
      ID: 1,
      Modified: '2018-03-15T10:52:10Z',
      Title: 'NewTitle'
    };
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$select=Id`) {
        return { Id: 'f64041f2-9818-4b67-92ff-3bc5dbbef27e' };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)`) {
        return listItemResponse;
      }

      throw 'Invalid request';
    });


    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">\n      <Actions>\n        \n    <Method Name="ParseAndSetFieldValue" Id="1" ObjectPathId="147">\n      <Parameters>\n        <Parameter Type="String">Title</Parameter>\n        <Parameter Type="String">NewTitle</Parameter>\n      </Parameters>\n    </Method>\n    <Method Name="ParseAndSetFieldValue" Id="2" ObjectPathId="147">\n      <Parameters>\n        <Parameter Type="String">customColumn</Parameter>\n        <Parameter Type="String">My custom column</Parameter>\n      </Parameters>\n    </Method>\n    <Method Name="ParseAndSetFieldValue" Id="3" ObjectPathId="147">\n      <Parameters>\n        <Parameter Type="String">ContentType</Parameter>\n        <Parameter Type="String">Item</Parameter>\n      </Parameters>\n    </Method>\n        <Method Name="SystemUpdate" Id="4" ObjectPathId="147" />\n      </Actions>\n      <ObjectPaths>\n        <Identity Id="147" Name="d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7:list:f64041f2-9818-4b67-92ff-3bc5dbbef27e:item:1,1" />\n      </ObjectPaths>\n    </Request>`) {
          return ']SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7618.1204","ErrorInfo":null,"TraceCorrelationId":"3e3e629e-f0e9-5000-9f31-c6758b453a4a"';
        }
      }

      throw 'Invalid request';
    });

    const actual = await spo.systemUpdateListItem(requestUrl, '1', logger, true, { Title: 'NewTitle', customColumn: 'My custom column' }, 'Item');
    assert.strictEqual(actual, listItemResponse);
  });

  it(`sets the content type of the list item with system update`, async () => {
    const listItemResponse = {
      Attachments: false,
      AuthorId: 3,
      ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
      Created: '2018-03-15T10:43:10Z',
      EditorId: 3,
      GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
      ID: 1,
      Modified: '2018-03-15T10:52:10Z',
      Title: 'NewTitle'
    };
    const listUrl = '/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$select=Id`) {
        return { Id: 'f64041f2-9818-4b67-92ff-3bc5dbbef27e' };
      }

      if (opts.url === `https://contoso.sharepoint.com/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)`) {
        return listItemResponse;
      }

      throw 'Invalid request';
    });


    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">\n      <Actions>\n        \n    <Method Name="ParseAndSetFieldValue" Id="1" ObjectPathId="147">\n      <Parameters>\n        <Parameter Type="String">ContentType</Parameter>\n        <Parameter Type="String">Item</Parameter>\n      </Parameters>\n    </Method>\n        <Method Name="SystemUpdate" Id="2" ObjectPathId="147" />\n      </Actions>\n      <ObjectPaths>\n        <Identity Id="147" Name="d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7:list:f64041f2-9818-4b67-92ff-3bc5dbbef27e:item:1,1" />\n      </ObjectPaths>\n    </Request>`) {
          return ']SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7618.1204","ErrorInfo":null,"TraceCorrelationId":"3e3e629e-f0e9-5000-9f31-c6758b453a4a"';
        }
      }

      throw 'Invalid request';
    });

    const actual = await spo.systemUpdateListItem(requestUrl, '1', logger, true, undefined, 'Item');
    assert.strictEqual(actual, listItemResponse);
  });

  it(`sets the list item without system update`, async () => {
    const listItemResponse = {
      Attachments: false,
      AuthorId: 3,
      ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
      Created: '2018-03-15T10:43:10Z',
      EditorId: 3,
      GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
      ID: 1,
      Modified: '2018-03-15T10:52:10Z',
      Title: 'NewTitle'
    };
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)`) {
        return listItemResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)/ValidateUpdateListItem()`) {
        return { value: [{ ItemId: 1 }] };
      }

      throw 'Invalid request';
    });

    const actual = await spo.updateListItem(requestUrl, '1');
    assert.strictEqual(actual, listItemResponse);
  });

  it(`sets the list item without system update and with options`, async () => {
    const listItemResponse = {
      Attachments: false,
      AuthorId: 3,
      ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
      Created: '2018-03-15T10:43:10Z',
      EditorId: 3,
      GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
      ID: 1,
      Modified: '2018-03-15T10:52:10Z',
      Title: 'NewTitle'
    };
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)`) {
        return listItemResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)/ValidateUpdateListItem()`) {
        return { value: [{ ItemId: 1 }] };
      }

      throw 'Invalid request';
    });

    const actual = await spo.updateListItem(requestUrl, '1', { Title: 'NewTitle', customColumn: 'My custom column' }, 'Item');
    assert.strictEqual(actual, listItemResponse);
  });

  it(`handles systemUpdate error when updating list item`, async () => {
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$select=Id`) {
        return { Id: 'f64041f2-9818-4b67-92ff-3bc5dbbef27e' };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ]);
        }

        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">\n      <Actions>\n        \n    <Method Name="ParseAndSetFieldValue" Id="1" ObjectPathId="147">\n      <Parameters>\n        <Parameter Type="String">Title</Parameter>\n        <Parameter Type="String">NewTitle</Parameter>\n      </Parameters>\n    </Method>\n        <Method Name="SystemUpdate" Id="2" ObjectPathId="147" />\n      </Actions>\n      <ObjectPaths>\n        <Identity Id="147" Name="d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7:list:f64041f2-9818-4b67-92ff-3bc5dbbef27e:item:1,1" />\n      </ObjectPaths>\n    </Request>`) {
          return 'ErrorMessage": "systemUpdate error"}';
        }
      }

      throw 'Invalid request';
    });

    try {
      await spo.systemUpdateListItem(requestUrl, '1', logger, true, { Title: 'NewTitle' });
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `Error occurred in systemUpdate operation - ErrorMessage": "systemUpdate error"}`);
    }
  });

  it(`handles no contenttype or properties error when updating list item`, async () => {
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')?$select=Id`) {
        return { Id: 'f64041f2-9818-4b67-92ff-3bc5dbbef27e' };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });

    try {
      await spo.systemUpdateListItem(requestUrl, '1', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `Either properties or contentTypeName must be provided for systemUpdateListItem.`);
    }
  });

  it(`handles error when a specific field fails when updating listitem`, async () => {
    const listUrl = '/sites/sales/lists/TestList';
    const requestUrl = `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')`;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetList('${formatting.encodeQueryParameter(listUrl)}')/items(1)/ValidateUpdateListItem()`) {
        return { value: [{ ErrorMessage: 'failed updating', 'FieldName': 'Title', 'HasException': true }] };
      }

      throw 'Invalid request';
    });

    try {
      await spo.updateListItem(requestUrl, '1');
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `Updating the items has failed with the following errors: ${os.EOL}- Title - failed updating`);
    }
  });

  it(`handles random error when requesting the ObjectIdentity fails`, async () => {
    const error = {
      ErrorInfo: {
        ErrorMessage: 'An unexpected error occured'
      }
    };

    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([error]);
        }
      }

      throw 'Invalid request';
    });

    try {
      await spo.requestObjectIdentity(webUrl, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, error.ErrorInfo.ErrorMessage);
    }
  });

  it(`handles ClientSvc unknown error when requesting the ObjectIdentity fails`, async () => {
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([{ "ErrorInfo": "error occurred" }]);
        }
      }

      throw 'Invalid request';
    });

    try {
      await spo.requestObjectIdentity(webUrl, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'ClientSvc unknown error');
    }
  });

  it(`handles error when _ObjectIdentity_ not found when requesting the ObjectIdentity fails`, async () => {
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestExpiresAt: new Date(),
      FormDigestTimeoutSeconds: 1800,
      WebFullUrl: 'https://contoso-admin.sharepoint.com'
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_vti_bin/client.svc/ProcessQuery`) {
        if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo": null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ]);
        }
      }

      throw 'Invalid request';
    });

    try {
      await spo.requestObjectIdentity(webUrl, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, 'Cannot proceed. _ObjectIdentity_ not found');
    }
  });

  it(`throws an error when waiting until a process is resulting in an error`, async () => {
    const objectIdentity = {
      "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation",
      "_ObjectIdentity_": "5492dba0-70ae-7000-66f6-1306e17a5220|908bed80-a04a-4433-b4a0-883d9847d110:1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4\nSpoOperation\nRemoveDeletedSite\n638306152161051712\nhttps%3a%2f%2fcontoso.sharepoint.com%2fteams%2fsales\nd8476b67-4a80-4261-a94f-431a2d0b5d3e",
      "IsComplete": true,
      "PollingInterval": 0
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7324.1200", "ErrorInfo": {
              "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4", "ErrorCode": -1, "ErrorTypeName": "SPException"
            }, "TraceCorrelationId": "b33c489e-009b-5000-8240-a8c28e5fd8b4"
          }
        ]);
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/contextinfo') {
        return { FormDigestValue: 'abc' };
      }

      throw 'Invalid request';
    });

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    try {
      await spo.waitUntilFinished({
        operationId: JSON.stringify(objectIdentity),
        siteUrl: 'https://contoso.sharepoint.com',
        logger,
        currentContext: ctx,
        verbose: false,
        debug: false
      });
      assert.fail('No error message thrown.');
    }
    catch (ex: any) {
      assert.deepStrictEqual(ex.message, 'An error has occurred.');
    }
  });

  it(`will retry when an operation is not finished after the first attempt`, async () => {
    let amountOfCalls = 0;
    const objectIdentity: SpoOperation = {
      _ObjectIdentity_: "5492dba0-70ae-7000-66f6-1306e17a5220|908bed80-a04a-4433-b4a0-883d9847d110:1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4\nSpoOperation\nRemoveDeletedSite\n638306152161051712\nhttps%3a%2f%2fcontoso.sharepoint.com%2fteams%2fsales\nd8476b67-4a80-4261-a94f-431a2d0b5d3e",
      IsComplete: false,
      PollingInterval: 0
    };
    sinon.stub(request, 'post').callsFake(async (opts) => {
      amountOfCalls++;
      if (opts.url === 'https://contoso.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        if (amountOfCalls > 2) {
          objectIdentity.IsComplete = true;
        }
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.24030.12011",
            "ErrorInfo": null,
            "TraceCorrelationId": "5492dba0-70ae-7000-66f6-1306e17a5220"
          },
          185,
          {
            "IsNull": false
          },
          186,
          objectIdentity
        ]);
      }

      if (opts.url === 'https://contoso.sharepoint.com/_api/contextinfo') {
        return { FormDigestValue: 'abc' };
      }

      throw 'Invalid request';
    });

    const ctx: FormDigestInfo = {
      FormDigestValue: 'value',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    };

    await spo.waitUntilFinished({
      operationId: JSON.stringify(objectIdentity),
      siteUrl: 'https://contoso.sharepoint.com',
      logger,
      currentContext: ctx,
      verbose: false,
      debug: false
    });
    assert.strictEqual(amountOfCalls, 4);
  });

  it('throws error when folder not found by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/web/GetFolderById('invalidFolderId')?$select=ServerRelativeUrl`) {
        throw `File Not Found`;
      }

      throw 'Invalid request';
    });

    try {
      await spo.getFolderServerRelativeUrl(webUrl, undefined, 'invalidFolderId', logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `File Not Found`);
    }
  });

  it('throws error when folder not found by url', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl='%2Fsites%2Fsales%2FinvalidFolderUrl')?$select=ServerRelativeUrl`) {
        throw `File Not Found`;
      }

      throw 'Invalid request';
    });

    try {
      await spo.getFolderServerRelativeUrl(webUrl, 'invalidFolderUrl', undefined, logger, true);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, `File Not Found`);
    }
  });

  it(`gets primary admin loginName from admin site`, async () => {
    const adminUrl = 'https://contoso-admin.sharepoint.com';
    const siteId = '0ead8b78-89e5-427f-b1bc-6e5a77ac191c';
    const primaryAdminLoginName = 'user1loginName';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_api/SPO.Tenant/sites('${siteId}')?$select=OwnerLoginName`) {
        return { OwnerLoginName: primaryAdminLoginName };
      }

      throw 'Invalid request';
    });

    const result = await spo.getPrimaryAdminLoginNameAsAdmin(adminUrl, siteId, logger, true);
    assert.strictEqual(result, primaryAdminLoginName);
  });

  it(`gets primary admin loginName from site`, async () => {
    const siteUrl = 'https://contoso.sharepoint.com';
    const primaryAdminLoginName = 'user1loginName';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${siteUrl}/_api/site/owner`) {
        return { LoginName: primaryAdminLoginName };
      }

      throw 'Invalid request';
    });

    const result = await spo.getPrimaryOwnerLoginFromSite(siteUrl, logger, true);
    assert.strictEqual(result, primaryAdminLoginName);
  });

  it(`retrieves a file with its properties sucessfully`, async () => {
    const id = 'b2307a39-e878-458b-bc90-03bc578531d6';
    const fileResponse = {
      ListItemAllFields: {
        FileSystemObjectType: 0,
        Id: 4,
        ServerRedirectedEmbedUri: 'https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
        ServerRedirectedEmbedUrl: 'https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview',
        ContentTypeId: '0x0101008E462E3ACE8DB844B3BEBF9473311889',
        ComplianceAssetId: null,
        Title: null,
        ID: 4,
        Created: '2018-02-05T09:42:36',
        AuthorId: 1,
        Modified: '2018-02-05T09:44:03',
        EditorId: 1,
        'OData__CopySource': null,
        CheckoutUserId: null,
        'OData__UIVersionString': '3.0',
        GUID: '2054f49e-0f76-46d4-ac55-50e1c057941c'
      },
      CheckInComment: '',
      CheckOutType: 2,
      ContentTag: '{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12',
      CustomizedPageStatus: 0,
      ETag: '\'{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\'',
      Exists: true,
      IrmEnabled: false,
      Length: '331673',
      Level: 1,
      LinkingUri: 'https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b',
      LinkingUrl: 'https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b',
      MajorVersion: 3,
      MinorVersion: 0,
      Name: 'Opendag maart 2018.docx',
      ServerRelativeUrl: '/sites/project-x/Documents/Test1.docx',
      TimeCreated: '2018-02-05T08:42:36Z',
      TimeLastModified: '2018-02-05T08:44:03Z',
      Title: '',
      UIVersion: 1536,
      UIVersionLabel: '3.0',
      UniqueId: 'b2307a39-e878-458b-bc90-03bc578531d6'
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/sales/_api/web/GetFileById('${formatting.encodeQueryParameter(id)}')`) {
        return fileResponse;
      }

      throw 'Invalid request';
    });

    const group = await spo.getFileById(webUrl, id, logger, true);
    assert.deepEqual(group, fileResponse);
  });

  it('correctly outputs result when calling createFileCopyJob', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const result = await spo.createFileCopyJob('https://contoso.sharepoint.com/sites/sales', 'https://contoso.sharepoint.com/sites/sales/Icons/Company.png', 'https://contoso.sharepoint.com/sites/marketing/Shared Documents');
    assert.deepStrictEqual(result, copyJobInfo);
  });

  it('correctly creates a copy job with default options when using createFileCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFileCopyJob('https://contoso.sharepoint.com/sites/sales', 'https://contoso.sharepoint.com/sites/sales/Icons/Company.png', 'https://contoso.sharepoint.com/sites/marketing/Shared Documents');
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons/Company.png'],
      options: {
        NameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Fail,
        AllowSchemaMismatch: true,
        BypassSharedLock: false,
        IgnoreVersionHistory: false,
        CustomizedItemName: undefined,
        IsMoveMode: false,
        IncludeItemPermissions: false,
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly creates a copy job with custom options when using createFileCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFileCopyJob(
      'https://contoso.sharepoint.com/sites/sales',
      'https://contoso.sharepoint.com/sites/sales/Icons/Company.png',
      'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Rename,
        bypassSharedLock: true,
        ignoreVersionHistory: true,
        newName: 'CompanyV2.png',
        operation: 'copy'
      }
    );
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons/Company.png'],
      options: {
        NameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Rename,
        AllowSchemaMismatch: true,
        BypassSharedLock: true,
        IgnoreVersionHistory: true,
        IsMoveMode: false,
        IncludeItemPermissions: false,
        CustomizedItemName: ['CompanyV2.png'],
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly creates a copy job with custom move options when using createFileCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFileCopyJob(
      'https://contoso.sharepoint.com/sites/sales',
      'https://contoso.sharepoint.com/sites/sales/Icons/Company.png',
      'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      {
        nameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Rename,
        bypassSharedLock: true,
        includeItemPermissions: true,
        newName: 'CompanyV2.png',
        operation: 'move'
      }
    );
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons/Company.png'],
      options: {
        NameConflictBehavior: CreateFileCopyJobsNameConflictBehavior.Rename,
        AllowSchemaMismatch: true,
        BypassSharedLock: true,
        IgnoreVersionHistory: false,
        IsMoveMode: true,
        IncludeItemPermissions: true,
        CustomizedItemName: ['CompanyV2.png'],
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly outputs result when calling createFolderCopyJob', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const result = await spo.createFolderCopyJob('https://contoso.sharepoint.com/sites/sales', 'https://contoso.sharepoint.com/sites/sales/Icons', 'https://contoso.sharepoint.com/sites/marketing/Shared Documents');
    assert.deepStrictEqual(result, copyJobInfo);
  });

  it('correctly creates a copy job with default options when using createFolderCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFolderCopyJob('https://contoso.sharepoint.com/sites/sales', 'https://contoso.sharepoint.com/sites/sales/Icons', 'https://contoso.sharepoint.com/sites/marketing/Shared Documents');
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons'],
      options: {
        NameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Fail,
        AllowSchemaMismatch: true,
        CustomizedItemName: undefined,
        IsMoveMode: false,
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly creates a copy job with custom options when using createFolderCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFolderCopyJob(
      'https://contoso.sharepoint.com/sites/sales',
      'https://contoso.sharepoint.com/sites/sales/Icons',
      'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        newName: 'Company icons',
        operation: 'copy'
      }
    );
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons'],
      options: {
        NameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        AllowSchemaMismatch: true,
        IsMoveMode: false,
        CustomizedItemName: ['Company icons'],
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly creates a copy job with custom move options when using createFolderCopyJob', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/CreateCopyJobs') {
        return {
          value: [
            copyJobInfo
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.createFolderCopyJob(
      'https://contoso.sharepoint.com/sites/sales',
      'https://contoso.sharepoint.com/sites/sales/Icons',
      'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      {
        nameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        newName: 'Company icons',
        operation: 'move'
      }
    );
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      destinationUri: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents',
      exportObjectUris: ['https://contoso.sharepoint.com/sites/sales/Icons'],
      options: {
        NameConflictBehavior: CreateFolderCopyJobsNameConflictBehavior.Rename,
        AllowSchemaMismatch: true,
        IsMoveMode: true,
        CustomizedItemName: ['Company icons'],
        SameWebCopyMoveOptimization: true
      }
    });
  });

  it('correctly polls for copy job status when using getCopyJobResult', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/GetCopyJobProgress') {
        if (postStub.callCount < 5) {
          return {
            JobState: 4,
            Logs: []
          };
        }

        if (postStub.callCount === 5) {
          return {
            JobState: 4,
            Logs: [
              JSON.stringify({
                Event: 'JobStart',
                JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
                Time: '08/10/2024 16:30:39.004',
                SiteId: '53dec431-9d4f-415b-b12b-010259d5b4e1',
                WebId: 'af102f32-b389-49dc-89bf-d116a17e0aa6',
                DBId: '5a926054-85d7-4cf6-85f0-c38fa01c4d39',
                FarmId: '823af112-cd95-49a2-adf5-eccb09c8ba5d',
                ServerId: 'a6145d7e-1b85-4124-895e-b1e618bfe5ae',
                SubscriptionId: '18c58817-3bc9-489d-ac63-f7264fb357e5',
                TotalRetryCount: '0',
                MigrationType: 'Copy',
                MigrationDirection: 'Import',
                CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
              }),
              JSON.stringify({
                Event: 'JobFinishedObjectInfo',
                JobId: '6d1eda82-0d1c-41eb-ab05-1d9cd4afe786',
                Time: '08/10/2024 18:59:40.145',
                SourceObjectFullUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Icons/Company.png',
                TargetServerUrl: 'https://contoso.sharepoint.com',
                TargetSiteId: '794dada8-4389-45ce-9559-0de74bf3554a',
                TargetWebId: '8de9b4d3-3c30-4fd0-a9d7-2452bd065555',
                TargetListId: '44b336a5-e397-4e22-a270-c39e9069b123',
                TargetObjectUniqueId: '15488d89-b82b-40be-958a-922b2ed79383',
                TargetObjectSiteRelativeUrl: 'Shared Documents/Icons/Company.png',
                CorrelationId: '5efd44a1-c034-9000-9692-4e1a1b3ca33b'
              })
            ]
          };
        }

        return {
          JobState: 0,
          Logs: [
            JSON.stringify({
              Event: 'JobEnd',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.008',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            })
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await spo.getCopyJobResult('https://contoso.sharepoint.com/sites/sales', copyJobInfo);

    const postRequests = postStub.getCalls();
    postRequests.forEach((request) =>
      assert.deepStrictEqual(request.args[0].data, { copyJobInfo: copyJobInfo })
    );
  });

  it('correctly returns result when using getCopyJobResult', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/GetCopyJobProgress') {
        return {
          JobState: 0,
          Logs: [
            JSON.stringify({
              Event: 'JobStart',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.004',
              SiteId: '53dec431-9d4f-415b-b12b-010259d5b4e1',
              WebId: 'af102f32-b389-49dc-89bf-d116a17e0aa6',
              DBId: '5a926054-85d7-4cf6-85f0-c38fa01c4d39',
              FarmId: '823af112-cd95-49a2-adf5-eccb09c8ba5d',
              ServerId: 'a6145d7e-1b85-4124-895e-b1e618bfe5ae',
              SubscriptionId: '18c58817-3bc9-489d-ac63-f7264fb357e5',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            }),
            JSON.stringify({
              Event: 'JobFinishedObjectInfo',
              JobId: '6d1eda82-0d1c-41eb-ab05-1d9cd4afe786',
              Time: '08/10/2024 18:59:40.145',
              SourceObjectFullUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Icons/Company.png',
              TargetServerUrl: 'https://contoso.sharepoint.com',
              TargetSiteId: '794dada8-4389-45ce-9559-0de74bf3554a',
              TargetWebId: '8de9b4d3-3c30-4fd0-a9d7-2452bd065555',
              TargetListId: '44b336a5-e397-4e22-a270-c39e9069b123',
              TargetObjectUniqueId: '15488d89-b82b-40be-958a-922b2ed79383',
              TargetObjectSiteRelativeUrl: 'Shared Documents/Icons/Company.png',
              CorrelationId: '5efd44a1-c034-9000-9692-4e1a1b3ca33b'
            }),
            JSON.stringify({
              Event: 'JobEnd',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.008',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            })
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    const result = await spo.getCopyJobResult('https://contoso.sharepoint.com/sites/sales', copyJobInfo);
    assert.deepStrictEqual(result,
      {
        Event: 'JobFinishedObjectInfo',
        JobId: '6d1eda82-0d1c-41eb-ab05-1d9cd4afe786',
        Time: '08/10/2024 18:59:40.145',
        SourceObjectFullUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared Documents/Icons/Company.png',
        TargetServerUrl: 'https://contoso.sharepoint.com',
        TargetSiteId: '794dada8-4389-45ce-9559-0de74bf3554a',
        TargetWebId: '8de9b4d3-3c30-4fd0-a9d7-2452bd065555',
        TargetListId: '44b336a5-e397-4e22-a270-c39e9069b123',
        TargetObjectUniqueId: '15488d89-b82b-40be-958a-922b2ed79383',
        TargetObjectSiteRelativeUrl: 'Shared Documents/Icons/Company.png',
        CorrelationId: '5efd44a1-c034-9000-9692-4e1a1b3ca33b'
      }
    );
  });

  it('correctly throws error when using getCopyJobResult', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso.sharepoint.com/sites/sales/_api/Site/GetCopyJobProgress') {
        return {
          JobState: 0,
          Logs: [
            JSON.stringify({
              Event: 'JobStart',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.004',
              SiteId: '53dec431-9d4f-415b-b12b-010259d5b4e1',
              WebId: 'af102f32-b389-49dc-89bf-d116a17e0aa6',
              DBId: '5a926054-85d7-4cf6-85f0-c38fa01c4d39',
              FarmId: '823af112-cd95-49a2-adf5-eccb09c8ba5d',
              ServerId: 'a6145d7e-1b85-4124-895e-b1e618bfe5ae',
              SubscriptionId: '18c58817-3bc9-489d-ac63-f7264fb357e5',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            }),
            JSON.stringify({
              Event: 'JobError',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.007',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              ObjectType: 'File',
              Url: 'Shared Documents/Icons/Company.png',
              Id: 'c194762b-3f54-4f5f-9f5c-eba26084e29d',
              SourceListItemIntId: '38',
              ErrorCode: '-2147024713',
              Message: 'A file or folder with the name Company.png already exists at the destination.',
              TargetListItemIntId: 'f9628bfc-1e80-4486-aa3e-25d1f1ac67f9',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            }),
            JSON.stringify({
              Event: 'JobEnd',
              JobId: 'fb4cc143-383c-4da0-bd91-02d2acbb01c7',
              Time: '08/10/2024 16:30:39.008',
              TotalRetryCount: '0',
              MigrationType: 'Copy',
              MigrationDirection: 'Import',
              CorrelationId: 'd8f444a1-10a8-9000-862c-0bad6eff1006'
            })
          ]
        };
      }

      throw 'Invalid request: ' + opts.url;
    });

    await assert.rejects(spo.getCopyJobResult('https://contoso.sharepoint.com/sites/sales', copyJobInfo),
      new Error('A file or folder with the name Company.png already exists at the destination.'));
  });

  it(`Gets site properties without included details as admin using provided url`, async () => {
    const siteId = 'b2307a39-e878-458b-bc90-03bc578531d6';
    const siteProperties = { SiteId: siteId };
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetSitePropertiesByUrl`) {
        return siteProperties;
      };

      throw 'Invalid request';
    });

    await spo.getSiteAdminPropertiesByUrl('https://contoso.sharepoint.com/sites/sales', false, logger, true);

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { url: 'https://contoso.sharepoint.com/sites/sales', includeDetail: false });
  });

  it(`Gets site properties with included details as admin using provided url`, async () => {
    const siteId = 'b2307a39-e878-458b-bc90-03bc578531d6';
    const siteProperties = { SiteId: siteId };
    sinon.stub(spo, 'getSpoAdminUrl').resolves('https://contoso-admin.sharepoint.com');
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetSitePropertiesByUrl`) {
        return siteProperties;
      };

      throw 'Invalid request';
    });

    await spo.getSiteAdminPropertiesByUrl('https://contoso.sharepoint.com/sites/sales', true, logger, true);

    assert.deepStrictEqual(postStub.firstCall.args[0].data, { url: 'https://contoso.sharepoint.com/sites/sales', includeDetail: true });
  });
});