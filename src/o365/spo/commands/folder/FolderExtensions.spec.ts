
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import * as sinon from 'sinon';
import { FolderExtensions } from './FolderExtensions'


describe('FolderExtensions', () => {

  let folderExtensions: FolderExtensions;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  let stubPostResponses: any = (
    folderAddResp = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {

      if (opts.url.indexOf('/_api/web/folders') > -1) {
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
    getFolderByServerRelativeUrlResp = null
  ) => {
    return sinon.stub(request, 'get').callsFake((opts) => {

      if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {

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

    folderExtensions.ensureFolder("abc", "abc", "abc")
    .then(res => {
      
      done('Sould reject, not resolve');
      
    },  (err:any) => {
      
      assert.equal(err, 'webFullUrl is not a valid URL');
      done();
    });
  });

  it('should reject if empty folder param', (done) => { 

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "", "abc")
    .then(res => {
      
      done('Sould reject, not resolve');

    },  (err:any) => {
      
      assert.equal(err, 'folderToEnsure cannot be empty');
      done();
    });
  });

  it('should reject if empty siteAccessToken param', (done) => { 

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "")
    .then(res => {
      
      done('Sould reject, not resolve');
      
    },  (err:any) => {
      
      assert.equal(err, 'siteAccessToken cannot be empty');
      done();
    });
  });

  it('should handle folder creation faliure', (done) => {
    
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Not found."}}}));
    });

    const expectedError = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Cannot create folder."}}});

    const folderCreationErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses(folderCreationErrorResp); 

    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      done('Should not resolve, but reject');
    },  (err:any) => {
      
      assert.equal(JSON.stringify(err), JSON.stringify(expectedError));
      done();
    });
  });

  it('should handle folder creation faliure (debug)', (done) => {
    
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Not found."}}}));
    });

    const expectedError = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Cannot create folder."}}});

    const folderCreationErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses(folderCreationErrorResp); 

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      done('Should not resolve, but reject');
    },  (err:any) => {
      
      assert.equal(JSON.stringify(err), JSON.stringify(expectedError));
      done();
    });
  });

  it('should succeed in adding folder if it does not exist (debug)', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Not found."}}}));
    });
    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses();
    
    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.lastCall.args[0], 'All sub-folders exist');
      done();

    },  (err:any) => {
      done(err);
    });
  });

  it('should succeed in adding folder if it does not exist', (done) => {
    const folderDoesNotExistErrorResp: any = new Promise<any>((resolve, reject) => {
      return reject(JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Not found."}}}));
    });
    stubGetResponses(folderDoesNotExistErrorResp);
    stubPostResponses();
    
    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.notCalled, true);
      done();

    },  (err:any) => {
      done(err);
    });
  });

  it('should succeed if all folders exist (debug)', (done) => {
    stubPostResponses();
    stubGetResponses();
    
    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.called, true);
      done();

    },  (err:any) => {
      done(err);
    });
  });

  it('should succeed if all folders exist', (done) => {
    stubPostResponses();
    stubGetResponses();

    folderExtensions = new FolderExtensions(cmdInstance, false);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "abc", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.called, false);
      done();

    },  (err:any) => {
      done(err);
    });
  });

  it('should remove end / from folder path', (done) => { 
    stubPostResponses();
    stubGetResponses();

    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "folder1/folder2/", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.calledWith('folder1/folder2'), true);
      done();
      
    },  (err:any) => {
      
      done(err);
    });
  });

  it('should remove end / from folder path', (done) => { 
    stubPostResponses();
    stubGetResponses();
    
    folderExtensions = new FolderExtensions(cmdInstance, true);

    folderExtensions.ensureFolder("https://contoso.sharepoint.com", "/folder2/folder3", "abc")
    .then(res => {
      
      assert.equal(cmdInstanceLogSpy.calledWith('folder2/folder3'), true);
      done();
      
    },  (err:any) => {
      
      done(err);
    });
  });
});