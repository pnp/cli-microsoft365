import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./file-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import { FolderExtensions } from '../folder/FolderExtensions';

describe(commands.FILE_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
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

      if (opts.url.indexOf('/common/oauth2/token') > -1) {

        return Promise.resolve('abc');

      } else if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {

        if (opts.url.indexOf('/CheckOut') > -1) {

          if(checkoutResp) {
            return checkoutResp;
          } else {
            return Promise.resolve({"odata.null":true});
          }

        } else if (opts.url.indexOf('Add') > -1) {

          if(fileAddResp) {
            return fileAddResp;
          } else {

            return Promise.resolve({"CheckInComment":"","CheckOutType":0,"ContentTag":"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428,159","CustomizedPageStatus":0,"ETag":"\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428\"","Exists":true,"IrmEnabled":false,"Length":"165114","Level":255,"LinkingUri":null,"LinkingUrl":"","MajorVersion":51,"MinorVersion":15,"Name":"MS365.jpg","ServerRelativeUrl":"/sites/VelinDev/Shared Documents/t1/MS365.jpg","TimeCreated":"2018-10-21T21:46:08Z","TimeLastModified":"2018-10-25T23:49:52Z","Title":"title4","UIVersion":26127,"UIVersionLabel":"51.15","UniqueId":"b0bc16bb-c8d9-4a24-bc04-fb52045f8bef"});
          }

        } else if (opts.url.indexOf('ValidateUpdateListItem') > -1) {

          if(validateUpdateListItemResp) {
            return validateUpdateListItemResp;
          } else {
            return Promise.resolve({"value":[{"ErrorMessage":null,"FieldName":"Title","FieldValue":"title4","HasException":false,"ItemId":212}]});
          }

        } else if (opts.url.indexOf('approve') > -1) {

          if (approveResp) {
            return approveResp;
          } else {
            return Promise.resolve({"odata.null":true});
          }
        } else if (opts.url.indexOf('publish') > -1) {

          if (publishResp) {
            return publishResp;
          } else {
            return Promise.resolve({"odata.null":true});
          }
        } else if (opts.url.indexOf('UndoCheckOut') > -1) {

          if(undoCheckOut) {
            return undoCheckOut;
          } else {
            return Promise.resolve({"odata.null":true});
          }
        } else if (opts.url.indexOf('CheckIn') > -1) {

          if(checkinResp) {
            return checkinResp;
          } else {
            return Promise.resolve({"odata.null":true});
          }

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

      if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if(opts.url.indexOf('ParentList') > -1){
          
          if(parentListResp) {
            return parentListResp;
          } else {
            return Promise.resolve({"EnableMinorVersions":true,"EnableModeration":false,"EnableVersioning":true,"Id":"0c7dc8ec-5871-4ac9-962c-f856102b917b"});
          }
          
        } else if (opts.url.indexOf('/Files') > -1) {

          if(getFileResp) {
            return getFileResp;
          } else {
            return Promise.resolve({"CheckInComment":"test checkin 33","CheckOutType":2,"ContentTag":"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409,152","CustomizedPageStatus":0,"ETag":"\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409\"","Exists":true,"IrmEnabled":false,"Length":"165114","Level":2,"LinkingUri":null,"LinkingUrl":"","MajorVersion":51,"MinorVersion":8,"Name":"MS365.jpg","ServerRelativeUrl":"/sites/VelinDev/Shared Documents/t1/MS365.jpg","TimeCreated":"2018-10-21T21:46:08Z","TimeLastModified":"2018-10-25T23:38:11Z","Title":"title4","UIVersion":26120,"UIVersionLabel":"51.8","UniqueId":"b0bc16bb-c8d9-4a24-bc04-fb52045f8bef"});
          }

        } else {

          if (getFolderByServerRelativeUrlResp) {
            return getFolderByServerRelativeUrlResp;
          } else {
            return Promise.resolve({"Exists":true,"IsWOPIEnabled":false,"ItemCount":1,"Name":"t1","ProgID":null,"ServerRelativeUrl":"/sites/VelinDev/Shared Documents/t1","TimeCreated":"2018-10-21T21:46:07Z","TimeLastModified":"2018-10-21T21:46:08Z","UniqueId":"b60f36ef-6425-4961-a515-327191b5ca8f","WelcomePage":""});
          }
        }
      } else if (opts.url.indexOf('contenttypes') > -1) {

        if(getContentTypesResp){
          return getContentTypesResp;
        } else {
          return Promise.resolve({ value: [{"Id":{"StringValue":"0x010100B8255567D591B64D8E99AB920B147A39"},"Name":"Document"},{"Id":{"StringValue":"0x0120001EE53A8A89A10E459930CBB9B7B596A1"},"Name":"Folder"},{"Id":{"StringValue":"0x01010200AE588D214ED1CF439DD4ED66926E5FB2"},"Name":"Picture"}] });
        }
      }
      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(fs, 'readFileSync').returns(new Buffer('abc'));
    ensureFolderStub = sinon.stub(FolderExtensions.prototype, 'ensureFolder').resolves();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });

  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      request.get,
      fs.existsSync
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      fs.readFileSync,
      fs.existsSync,
      FolderExtensions.prototype.ensureFolder
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name, commands.FILE_ADD);
  });

  it('allows unknown options', () => {
    const actual = command.allowUnknownOptions();
    assert.equal(actual, true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.FILE_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call ensure folder when folder not found', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Not Found."}}});
    const getFolderByServerRelativeUrlResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(getFolderByServerRelativeUrlResp);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(ensureFolderStub.called, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should proceed with no error if file does not exist in the folder', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File not found."}}});
    const fileNotFoundResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, fileNotFoundResp);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }}, () => {

      try {
        assert.equal(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle checkout error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Checkout Error."}}});
    const checkoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(checkoutResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle file add error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File add error."}}});
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, fileAddResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle get list response error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: List does not exist."}}});
    const listResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, null, listResp);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle content type response error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: ContentType does not exist."}}});
    const contentTypeResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses();
    stubGetResponses(null, null, null, contentTypeResp);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folderServerRelativePath,
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(cmdInstanceLogSpy.calledWith(`folder path: ${folderServerRelativePath}...`), true);
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Specified content type \'abc\' doesn\'t exist on the target list')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle list item update response error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Item update error."}}});
    const validateUpdateListItemResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, validateUpdateListItemResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle list item field value update response error', (done) => {
    const expectedResult: any = {"value":[{"ErrorMessage":null,"FieldName":"Title","FieldValue":"fsd","HasException":false,"ItemId":120},{"ErrorMessage":"check in comment x","FieldName":"_CheckinComment","FieldValue":"check in comment x","HasException":true,"ItemId":120}]};
    const validateUpdateListItemResp: any = new Promise<any>((resolve, reject) => {
      return resolve(expectedResult);
    });
    stubPostResponses(null, null, validateUpdateListItemResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }}, (err: any) => {

      try {
        const error: string = `Update field value error: ${JSON.stringify(expectedResult.value)}`;
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(error)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle file checkin error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Checkin error."}}});
    const checkinResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, null, null, null, checkinResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        checkInComment: 'abc',
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle approve list item response error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Approve error."}}});
    const aproveResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, aproveResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true,
        verbose: true,
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle publish list item response error', (done) => {
    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Publish error."}}});
    const publishResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });
    stubPostResponses(null, null, null, null, publishResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true,
        verbose: true,
        debug: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error when --publish used, but list moderation and minor ver enabled', (done) => {
    const listSettingsResp: any = new Promise<any>((resolve, reject) => {
      return resolve({"EnableMinorVersions":true,"EnableModeration":true,"EnableVersioning":true,"Id":"0c7dc8ec-5871-4ac9-962c-f856102b917b"})
    });
    
    stubPostResponses();
    stubGetResponses(null, null, listSettingsResp);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true
      }}, (err: any) => {

      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('The file cannot be published without approval. Moderation for this list is enabled. Use the --approve option instead of --publish to approve and publish the file')));
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(cmdInstanceLogSpy.lastCall.args[0], 'DONE');
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(cmdInstanceLogSpy.lastCall.args[0], 'DONE');
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should succeed approve', (done) => {
    stubPostResponses();
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true
      }}, (err: any) => {

      try {
        assert.equal(cmdInstanceLogSpy.notCalled, true);
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

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }}, (err: any) => {

      try {
        assert.equal(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error if cannot rollback checkout (verbose)', (done) => {

    const expectedFileAddError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File add error."}}});
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedFileAddError);
    });

    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Checkout Error."}}});
    const rollbackCheckoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubPostResponses(null, fileAddResp, null, null, null, rollbackCheckoutResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true,
        verbose: true
      }}, (err: any) => {

      try {
        //assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should error if cannot rollback checkout', (done) => {

    const expectedFileAddError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: File add error."}}});
    const fileAddResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedFileAddError);
    });

    const expectedError: any = JSON.stringify({"odata.error":{"code":"-2130575338, Microsoft.SharePoint.SPException","message":{"lang":"en-US","value":"Error: Checkout Error."}}});
    const rollbackCheckoutResp: any = new Promise<any>((resolve, reject) => {
      return reject(expectedError);
    });

    stubPostResponses(null, fileAddResp, null, null, null, rollbackCheckoutResp);
    stubGetResponses();

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
      }}, (err: any) => {

      try {
        //assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(expectedFileAddError)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option not valid url', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'abc'} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the path option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', folder: 'abc'} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the folder option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', path: 'abc'} });
    assert.notEqual(actual, true);
  });


  it('fails validation if the wrong path option specified', () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com',
      folder: 'abc',
      path: 'abc'
    } });
    assert.notEqual(actual, true);
  });

  it('fails validation if --approveComment specified, but not --approve', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com',
      folder: 'abc',
      path: 'abc',
      approveComment: 'abc'
    } });
    assert.notEqual(actual, true);
  });

  it('fails validation if --publishComment specified, but not --publish', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com',
      folder: 'abc',
      path: 'abc',
      publishComment: 'abc'
    } });
    assert.notEqual(actual, true);
  });

  it('passed validation if options correct', () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = (command.validate() as CommandValidate)({ options: { 
      webUrl: 'https://contoso.sharepoint.com',
      folder: 'abc',
      path: 'abc'
    } });
    assert.equal(actual, true);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.FILE_ADD));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});