import * as assert from 'assert';
import request from '../../request';
import Utils from '../../Utils';
import * as sinon from 'sinon';
import { FolderExtensions } from './FolderExtensions'

describe('FolderExtensions', () => {
  let folderExtensions: FolderExtensions;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let stubPostResponses: any = (
    folderAddResp: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativePath') > -1) {
        if (folderAddResp) {
          return folderAddResp;
        } else {
          return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "4t4", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/4t4", "TimeCreated": "2018-10-26T22:50:27Z", "TimeLastModified": "2018-10-26T22:50:27Z", "UniqueId": "3f5428e2-b0a8-4d35-87df-89621ed5b457", "WelcomePage": "" });
        }

      }
      return Promise.reject('Invalid request');
    });
  }

  let stubGetResponses: any = (
    getFolderByServerRelativeUrlResp: any = null
  ) => {
    return sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if (getFolderByServerRelativeUrlResp) {
          return getFolderByServerRelativeUrlResp;
        } else {
          return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "f", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/4t4/f", "TimeCreated": "2018-10-26T22:54:19Z", "TimeLastModified": "2018-10-26T22:54:20Z", "UniqueId": "0d680f20-53da-4516-b3f6-ed98b1d928e8", "WelcomePage": "" });
        }
      }
      return Promise.reject('Invalid request');
    });
  }

  beforeEach(() => {
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get,
      cmdInstance.log
    ]);
  });

  it('should reject if wrong url param', (done) => {
    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("abc", "abc")
      .then(res => {

        done('Should reject, not resolve');

      }, (err: any) => {

        assert.strictEqual(err, 'webFullUrl is not a valid URL');
        done();
      });
  });

  it('should reject if empty folder param', (done) => {

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "")
      .then(res => {

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

    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {

        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0], 'All sub-folders exist');
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

    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {

        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();

      }, (err: any) => {
        done(err);
      });
  });

  it('should succeed if all folders exist (debug)', (done) => {
    stubPostResponses();
    stubGetResponses();

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {

        assert.strictEqual(cmdInstanceLogSpy.called, true);
        done();

      }, (err: any) => {
        done(err);
      });
  });

  it('should succeed if all folders exist', (done) => {
    stubPostResponses();
    stubGetResponses();

    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc")
      .then(res => {

        assert.strictEqual(cmdInstanceLogSpy.called, false);
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "/folder2/folder3")
      .then(res => {

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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com/sites/Site1/", "/folder2/folder3/")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com/sites/site1/", "PnP1/Folder2/")
      .then(res => {
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

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com/sites/Site1", "/folder2/folder3")
      .then(res => {
        assert.strictEqual(getStubs.getCall(0).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2FSite1%2Ffolder2\')');
        assert.strictEqual(getStubs.getCall(1).args[0].url, 'https://contoso.sharepoint.com/sites/Site1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2FSite1%2Ffolder2%2Ffolder3\')');
        done();
      }, (err: any) => {

        done(err);
      });
  });
});
