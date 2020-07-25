import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./file-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import { FolderExtensions } from '../../FolderExtensions';

describe(commands.FILE_ADD, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let ensureFolderStub: sinon.SinonStub;

  let stubPostResponses: any = (
    checkoutResp: any = null,
    fileAddResp: any = null,
    validateUpdateListItemResp: any = null,
    approveResp: any = null,
    publishResp: any = null,
    undoCheckOut: any = null,
    checkinResp: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/CheckOut') > -1) {

          if (checkoutResp) {
            return checkoutResp;
          } else {
            return Promise.resolve({ "odata.null": true });
          }

        } else if ((opts.url as string).indexOf('Add') > -1) {

          if (fileAddResp) {
            return fileAddResp;
          } else {

            return Promise.resolve({ "CheckInComment": "", "CheckOutType": 0, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428,159", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 255, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 15, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:49:52Z", "Title": "title4", "UIVersion": 26127, "UIVersionLabel": "51.15", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" });
          }

        } else if ((opts.url as string).indexOf('ValidateUpdateListItem') > -1) {

          if (validateUpdateListItemResp) {
            return validateUpdateListItemResp;
          } else {
            return Promise.resolve({ "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "title4", "HasException": false, "ItemId": 212 }] });
          }

        } else if ((opts.url as string).indexOf('approve') > -1) {

          if (approveResp) {
            return approveResp;
          } else {
            return Promise.resolve({ "odata.null": true });
          }
        } else if ((opts.url as string).indexOf('publish') > -1) {

          if (publishResp) {
            return publishResp;
          } else {
            return Promise.resolve({ "odata.null": true });
          }
        } else if ((opts.url as string).indexOf('UndoCheckOut') > -1) {

          if (undoCheckOut) {
            return undoCheckOut;
          } else {
            return Promise.resolve({ "odata.null": true });
          }
        } else if ((opts.url as string).indexOf('CheckIn') > -1) {

          if (checkinResp) {
            return checkinResp;
          } else {
            return Promise.resolve({ "odata.null": true });
          }

        } else if ((opts.url as string).indexOf('/StartUpload') !== -1) {

          return Promise.resolve({ "d": { "StartUpload": "0" } });

        } else if ((opts.url as string).indexOf('/cancelupload') !== -1) {

          return Promise.resolve({ "d": { "CancelUpload": null } });

        } else if ((opts.url as string).indexOf('/ContinueUpload') !== -1) {

          return Promise.resolve({ "d": { "ContinueUpload": "262144000" } });

        } else if ((opts.url as string).indexOf('/FinishUpload') !== -1) {

          return Promise.resolve({ "d": { "__metadata": { "id": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared Documents/IMG_9977.zip')", "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')", "type": "SP.File" }, "Author": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Author" } }, "CheckedOutByUser": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/CheckedOutByUser" } }, "EffectiveInformationRightsManagementSettings": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/EffectiveInformationRightsManagementSettings" } }, "InformationRightsManagementSettings": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/InformationRightsManagementSettings" } }, "ListItemAllFields": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/ListItemAllFields" } }, "LockedByUser": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/LockedByUser" } }, "ModifiedBy": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/ModifiedBy" } }, "Properties": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Properties" } }, "VersionEvents": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/VersionEvents" } }, "Versions": { "__deferred": { "uri": "https://velingeorgiev.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Versions" } }, "CheckInComment": "", "CheckOutType": 2, "ContentTag": "{1CDD37BD-BC3E-41DD-AB6C-89E3E975EEEB},2,2", "CustomizedPageStatus": 0, "ETag": "\"{1CDD37BD-BC3E-41DD-AB6C-89E3E975EEEB},2\"", "Exists": true, "IrmEnabled": false, "Length": "638194380", "Level": 1, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 1, "MinorVersion": 0, "Name": "IMG_9977.zip", "ServerRelativeUrl": "/Shared Documents/IMG_9977.zip", "TimeCreated": "2020-01-21T12:30:16Z", "TimeLastModified": "2020-01-21T12:32:18Z", "Title": null, "UIVersion": 512, "UIVersionLabel": "1.0", "UniqueId": "1cdd37bd-bc3e-41dd-ab6c-89e3e975eeeb" } });
        }
      }
      return Promise.reject('Invalid request');
    });
  }

  let stubGetResponses: any = (
    getFolderByServerRelativeUrlResp: any = null,
    getFileResp: any = null,
    parentListResp: any = null,
    getContentTypesResp: any = null
  ) => {
    return sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('ParentList') > -1) {

          if (parentListResp) {
            return parentListResp;
          } else {
            return Promise.resolve({ "EnableMinorVersions": true, "EnableModeration": false, "EnableVersioning": true, "Id": "0c7dc8ec-5871-4ac9-962c-f856102b917b" });
          }

        } else if ((opts.url as string).indexOf('/Files') > -1) {

          if (getFileResp) {
            return getFileResp;
          } else {
            return Promise.resolve({ "CheckInComment": "test checkin 33", "CheckOutType": 2, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409,152", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 2, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 8, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:38:11Z", "Title": "title4", "UIVersion": 26120, "UIVersionLabel": "51.8", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" });
          }

        } else {

          if (getFolderByServerRelativeUrlResp) {
            return getFolderByServerRelativeUrlResp;
          } else {
            return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "t1", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1", "TimeCreated": "2018-10-21T21:46:07Z", "TimeLastModified": "2018-10-21T21:46:08Z", "UniqueId": "b60f36ef-6425-4961-a515-327191b5ca8f", "WelcomePage": "" });
          }
        }
      } else if ((opts.url as string).indexOf('contenttypes') > -1) {

        if (getContentTypesResp) {
          return getContentTypesResp;
        } else {
          return Promise.resolve({ value: [{ "Id": { "StringValue": "0x010100B8255567D591B64D8E99AB920B147A39" }, "Name": "Document" }, { "Id": { "StringValue": "0x0120001EE53A8A89A10E459930CBB9B7B596A1" }, "Name": "Folder" }, { "Id": { "StringValue": "0x01010200AE588D214ED1CF439DD4ED66926E5FB2" }, "Name": "Picture" }] });
        }
      }
      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(fs, 'readFileSync').returns(Buffer.from('abc'));
    ensureFolderStub = sinon.stub(FolderExtensions.prototype, 'ensureFolder').resolves();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    sinon.stub(fs, 'statSync').returns({ size: 1234 } as any);
    sinon.stub(fs, 'openSync').returns(3);
    sinon.stub(fs, 'readSync').returns(10485760);
    sinon.stub(fs, 'closeSync');
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get,
      fs.existsSync,
      fs.statSync,
      fs.openSync,
      fs.readSync,
      fs.closeSync
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      fs.readFileSync,
      fs.existsSync,
      FolderExtensions.prototype.ensureFolder,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.FILE_ADD);
  });

  it('allows unknown options', () => {
    const actual = command.allowUnknownOptions();
    assert.strictEqual(actual, true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should call ensure folder when folder not found', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not Found." } } });
    const getFolderByServerRelativeUrlResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(getFolderByServerRelativeUrlResp);

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, () => {
      try {
        assert.strictEqual(ensureFolderStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should proceed with no error if file does not exist in the folder', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File not found." } } });
    const fileNotFoundResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, fileNotFoundResp);

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle checkout error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });
    const checkoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(checkoutResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle file add error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, fileAddResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle get list response error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: List does not exist." } } });
    const listResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, null, listResp);

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle content type response error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: ContentType does not exist." } } });
    const contentTypeResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, null, null, contentTypeResp);

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should resolve server relative url specified for the folder option', (done) => {
    stubPostResponses();
    stubGetResponses();

    const folderServerRelativePath: string = '/sites/project-x/Shared%20Documents/t1';

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folderServerRelativePath,
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(cmdInstanceLogSpy.calledWith(`folder path: ${folderServerRelativePath}...`), true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should resolve safe filename when path (bash) contains apostrophes in folders and filename', (done) => {
    stubPostResponses();
    stubGetResponses();

    const unsafePath: string = '/Users/user/Projects/TEST\'FOLDER/TEST\'FILE.txt';

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: unsafePath,
        contentType: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(cmdInstanceLogSpy.calledWith(`file name: TEST''FILE.txt...`), true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle non existing content type', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Specified content type \'abc\' doesn\'t exist on the target list')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle list item update response error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Item update error." } } });
    const validateUpdateListItemResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, validateUpdateListItemResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle list item field value update response error', (done) => {
    const expectedResult: any = { "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "fsd", "HasException": false, "ItemId": 120 }, { "ErrorMessage": "check in comment x", "FieldName": "_CheckinComment", "FieldValue": "check in comment x", "HasException": true, "ItemId": 120 }] };
    const validateUpdateListItemResp: any = new Promise<any>((resolve, reject) => {
      return resolve(expectedResult);
    });
    stubPostResponses(null, null, validateUpdateListItemResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }
    }, (err: any) => {
      try {
        const error: string = `Update field value error: ${JSON.stringify(expectedResult.value)}`;
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle file checkin error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkin error." } } });
    const checkinResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, null, null, null, checkinResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        checkInComment: 'abc',
        debug: true
      }
    }, (err: any) => {

      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle approve list item response error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Approve error." } } });
    const aproveResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, aproveResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true,
        verbose: true,
        debug: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle publish list item response error', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Publish error." } } });
    const publishResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, null, publishResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true,
        verbose: true,
        debug: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error when --publish used, but list moderation and minor ver enabled', (done) => {
    const listSettingsResp: any = new Promise<any>((resolve, reject) => {
      return resolve({ "EnableMinorVersions": true, "EnableModeration": true, "EnableVersioning": true, "Id": "0c7dc8ec-5871-4ac9-962c-f856102b917b" })
    });

    stubPostResponses();
    stubGetResponses(null, null, listSettingsResp);

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The file cannot be published without approval. Moderation for this list is enabled. Use the --approve option instead of --publish to approve and publish the file')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call "DONE" when in verbose', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0], 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores global options when creating request body', (done) => {
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        publish: true,
        debug: true,
        verbose: true,
        output: "text"
      }
    }, () => {
      try {
        assert.deepEqual(postRequests.secondCall.args[0].body, {
          bNewDocumentUpdate: true,
          checkInComment: '',
          formValues: [{ FieldName: 'Title', FieldValue: 'abc' }, { FieldName: 'ContentType', FieldValue: 'Picture' }]
        });
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should perform single request upload for file up to 250 MB', (done) => {
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    Utils.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 250 * 1024 * 1024 } as any); // 250 MB

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, () => {
      try {
        assert.notStrictEqual(postRequests.lastCall.args[0].url.indexOf(`/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%2520Documents%2Ft1')/Files/Add`), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should perform chunck upload on files over 250 MB (debug)', (done) => {
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    Utils.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, () => {
      try {
        assert.notStrictEqual(postRequests.firstCall.args[0].url.indexOf('/StartUpload'), -1);
        assert.notStrictEqual(postRequests.getCalls()[2].args[0].url.indexOf('/ContinueUpload'), -1);
        assert.notStrictEqual(postRequests.lastCall.args[0].url.indexOf('/FinishUpload'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should cancel chunck upload on files over 250 MB on error', (done) => {
    stubGetResponses();
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/StartUpload') !== -1) {

          return Promise.resolve({ "d": { "StartUpload": "0" } });

        } else if ((opts.url as string).indexOf('/cancelupload') !== -1) {

          return Promise.resolve({ "d": { "CancelUpload": null } });

        } else if ((opts.url as string).indexOf('/ContinueUpload') !== -1) {

          return Promise.reject({ "error": "123" });

        }
      }
      return Promise.reject('Invalid request');
    });

    Utils.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, '123');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle fail to read file error', (done) => {
    stubGetResponses();
    stubPostResponses();

    Utils.restore([fs.statSync, fs.openSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB
    sinon.stub(fs, 'openSync').throws(new Error('openSync error'));

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'openSync error');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should try closeSync on error', (done) => {
    stubGetResponses();
    stubPostResponses();

    Utils.restore([fs.statSync, fs.openSync, , fs.readSync, , fs.closeSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB
    sinon.stub(fs, 'openSync').returns(3);
    sinon.stub(fs, 'readSync').throws(new Error('readSync error'));
    sinon.stub(fs, 'closeSync').throws(new Error('failed to closeSync'));

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'readSync error');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed updating list item metadata (verbose)', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        publish: true,
        verbose: true,
        debug: true
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0], 'DONE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed updating list item metadata', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        publish: true
      }
    }, () => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets field with the same name as a command option but different casing', (done) => {
    stubGetResponses();
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/CheckOut') > -1) {
          return Promise.resolve({ "odata.null": true });
        }
        else if ((opts.url as string).indexOf('Add') > -1) {
          return Promise.resolve({ "CheckInComment": "", "CheckOutType": 0, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428,159", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 255, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 15, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:49:52Z", "Title": "title4", "UIVersion": 26127, "UIVersionLabel": "51.15", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" });
        }
        else if ((opts.url as string).indexOf('ValidateUpdateListItem') > -1) {
          if (opts.body.formValues.filter((f: any) => f.FieldName === 'Folder').length > 0) {
            return Promise.resolve({ "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "title4", "HasException": false, "ItemId": 212 }] });
          }
          else {
            return Promise.reject('Field Folder missing');
          }
        }
        else if ((opts.url as string).indexOf('approve') > -1) {
          return Promise.resolve({ "odata.null": true });
        }
        else if ((opts.url as string).indexOf('publish') > -1) {
          return Promise.resolve({ "odata.null": true });
        }
        else if ((opts.url as string).indexOf('UndoCheckOut') > -1) {
          return Promise.resolve({ "odata.null": true });
        }
        else if ((opts.url as string).indexOf('CheckIn') > -1) {
          return Promise.resolve({ "odata.null": true });
        }
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        Folder: 'Folder',
        publish: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(new Error(err.message));
      }
    });
  });

  it('should succeed approve', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed when with checkout option', (done) => {
    stubPostResponses();
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error if cannot rollback checkout (verbose)', (done) => {
    const expectedFileAddError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedFileAddError);
    });

    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });
    const rollbackCheckoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubPostResponses(null, fileAddResp, null, null, null, rollbackCheckoutResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true,
        verbose: true
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error if cannot rollback checkout', (done) => {
    const expectedFileAddError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedFileAddError);
    });

    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });
    const rollbackCheckoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubPostResponses(null, fileAddResp, null, null, null, rollbackCheckoutResp);
    stubGetResponses();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the webUrl option not valid url', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the wrong path option specified', () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if --approveComment specified, but not --approve', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc',
        approveComment: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if --publishComment specified, but not --publish', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc',
        publishComment: 'abc'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if options correct', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });
});