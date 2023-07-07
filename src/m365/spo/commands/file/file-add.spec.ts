import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./file-add');

describe(commands.FILE_ADD, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let ensureFolderStub: sinon.SinonStub;

  const stubPostResponses: any = (
    checkoutResp: any = null,
    fileAddResp: any = null,
    validateUpdateListItemResp: any = null,
    approveResp: any = null,
    publishResp: any = null,
    undoCheckOut: any = null,
    checkinResp: any = null,
    validateUpdateListItemRespPass: any = null
  ) => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/CheckOut') > -1) {

          if (checkoutResp) {
            throw checkoutResp;
          }
          else {
            return { "odata.null": true };
          }

        }
        else if ((opts.url as string).indexOf('Add') > -1) {

          if (fileAddResp) {
            throw fileAddResp;
          }
          else {

            return { "CheckInComment": "", "CheckOutType": 0, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428,159", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 255, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 15, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:49:52Z", "Title": "title4", "UIVersion": 26127, "UIVersionLabel": "51.15", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" };
          }

        }
        else if ((opts.url as string).indexOf('ValidateUpdateListItem') > -1) {

          if (validateUpdateListItemResp) {
            throw validateUpdateListItemResp;
          }
          else if (validateUpdateListItemRespPass) {
            return validateUpdateListItemRespPass;
          }
          else {
            return { "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "title4", "HasException": false, "ItemId": 212 }] };
          }

        }
        else if ((opts.url as string).indexOf('approve') > -1) {

          if (approveResp) {
            throw approveResp;
          }
          else {
            return { "odata.null": true };
          }
        }
        else if ((opts.url as string).indexOf('publish') > -1) {

          if (publishResp) {
            throw publishResp;
          }
          else {
            return { "odata.null": true };
          }
        }
        else if ((opts.url as string).indexOf('UndoCheckOut') > -1) {

          if (undoCheckOut) {
            throw undoCheckOut;
          }
          else {
            return { "odata.null": true };
          }
        }
        else if ((opts.url as string).indexOf('CheckIn') > -1) {

          if (checkinResp) {
            throw checkinResp;
          }
          else {
            return { "odata.null": true };
          }

        }
        else if ((opts.url as string).indexOf('/StartUpload') !== -1) {

          return { "d": { "StartUpload": "0" } };

        }
        else if ((opts.url as string).indexOf('/cancelupload') !== -1) {

          return { "d": { "CancelUpload": null } };

        }
        else if ((opts.url as string).indexOf('/ContinueUpload') !== -1) {

          return { "d": { "ContinueUpload": "262144000" } };

        }
        else if ((opts.url as string).indexOf('/FinishUpload') !== -1) {

          return { "d": { "__metadata": { "id": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared Documents/IMG_9977.zip')", "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')", "type": "SP.File" }, "Author": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Author" } }, "CheckedOutByUser": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/CheckedOutByUser" } }, "EffectiveInformationRightsManagementSettings": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/EffectiveInformationRightsManagementSettings" } }, "InformationRightsManagementSettings": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/InformationRightsManagementSettings" } }, "ListItemAllFields": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/ListItemAllFields" } }, "LockedByUser": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/LockedByUser" } }, "ModifiedBy": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/ModifiedBy" } }, "Properties": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Properties" } }, "VersionEvents": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/VersionEvents" } }, "Versions": { "__deferred": { "uri": "https://contoso.sharepoint.com/_api/Web/GetFileByServerRelativePath(decodedurl='/Shared%20Documents/IMG_9977.zip')/Versions" } }, "CheckInComment": "", "CheckOutType": 2, "ContentTag": "{1CDD37BD-BC3E-41DD-AB6C-89E3E975EEEB},2,2", "CustomizedPageStatus": 0, "ETag": "\"{1CDD37BD-BC3E-41DD-AB6C-89E3E975EEEB},2\"", "Exists": true, "IrmEnabled": false, "Length": "638194380", "Level": 1, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 1, "MinorVersion": 0, "Name": "IMG_9977.zip", "ServerRelativeUrl": "/Shared Documents/IMG_9977.zip", "TimeCreated": "2020-01-21T12:30:16Z", "TimeLastModified": "2020-01-21T12:32:18Z", "Title": null, "UIVersion": 512, "UIVersionLabel": "1.0", "UniqueId": "1cdd37bd-bc3e-41dd-ab6c-89e3e975eeeb" } };
        }
      }
      throw 'Invalid request';
    });
  };

  const stubGetResponses: any = (
    getFolderByServerRelativeUrlResp: any = null,
    getFileResp: any = null,
    parentListResp: any = null,
    getContentTypesResp: any = null,
    parentListRespSuccess: any = null
  ) => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {

      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('ParentList') > -1) {

          if (parentListResp) {
            throw parentListResp;
          }
          else if (parentListRespSuccess) {
            return parentListRespSuccess;
          }
          else {
            return { "EnableMinorVersions": true, "EnableModeration": false, "EnableVersioning": true, "Id": "0c7dc8ec-5871-4ac9-962c-f856102b917b" };
          }

        }
        else if ((opts.url as string).indexOf('/Files') > -1) {

          if (getFileResp) {
            throw getFileResp;
          }
          else {
            return { "CheckInComment": "test checkin 33", "CheckOutType": 2, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409,152", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},409\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 2, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 8, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:38:11Z", "Title": "title4", "UIVersion": 26120, "UIVersionLabel": "51.8", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" };
          }

        }
        else {

          if (getFolderByServerRelativeUrlResp) {
            throw getFolderByServerRelativeUrlResp;
          }
          else {
            return { "Exists": true, "IsWOPIEnabled": false, "ItemCount": 1, "Name": "t1", "ProgID": null, "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1", "TimeCreated": "2018-10-21T21:46:07Z", "TimeLastModified": "2018-10-21T21:46:08Z", "UniqueId": "b60f36ef-6425-4961-a515-327191b5ca8f", "WelcomePage": "" };
          }
        }
      }
      else if ((opts.url as string).indexOf('contenttypes') > -1) {

        if (getContentTypesResp) {
          throw getContentTypesResp;
        }
        else {
          return { value: [{ "Id": { "StringValue": "0x010100B8255567D591B64D8E99AB920B147A39" }, "Name": "Document" }, { "Id": { "StringValue": "0x0120001EE53A8A89A10E459930CBB9B7B596A1" }, "Name": "Folder" }, { "Id": { "StringValue": "0x01010200AE588D214ED1CF439DD4ED66926E5FB2" }, "Name": "Picture" }] };
        }
      }
      throw 'Invalid request';
    });
  };

  const stubFs = () => {
    sinon.stub(fs, 'readFileSync').returns(Buffer.from('abc'));
    sinon.stub(fs, 'statSync').returns({ size: 1234 } as any);
    sinon.stub(fs, 'openSync').returns(3);
    sinon.stub(fs, 'readSync').returns(10485760);
    sinon.stub(fs, 'closeSync');
  };

  before(() => {
    ensureFolderStub = sinon.stub(spo, 'ensureFolder').resolves();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(Buffer, 'alloc').returns(Buffer.from('abc'));
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      request.get,
      fs.existsSync,
      fs.statSync,
      fs.openSync,
      fs.readSync,
      fs.closeSync,
      fs.readFileSync
    ]);
  });

  after(() => {
    sinon.restore();
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

  it('should call ensure folder when folder not found', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Not Found." } } });

    stubFs();
    stubPostResponses();
    stubGetResponses(expectedError);

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    });
    assert.strictEqual(ensureFolderStub.called, true);
  });

  it('should proceed with no error if file does not exist in the folder', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File not found." } } });

    stubPostResponses();
    stubGetResponses(null, expectedError);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }
    }));
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('should handle checkout error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });

    stubPostResponses(expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle file add error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });

    stubFs();
    stubPostResponses(null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle get list response error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: List does not exist." } } });

    stubFs();
    stubPostResponses();
    stubGetResponses(null, null, expectedError);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle content type response error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: ContentType does not exist." } } });

    stubFs();
    stubPostResponses();
    stubGetResponses(null, null, null, expectedError);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should resolve server relative url specified for the folder option', async () => {
    stubPostResponses();
    stubGetResponses();

    const folderServerRelativePath: string = '/sites/project-x/Shared%20Documents/t1';

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: folderServerRelativePath,
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }));
    assert.strictEqual(loggerLogToStderrSpy.calledWith(`folder path: ${folderServerRelativePath}...`), true);
  });

  it('should resolve safe filename when path (bash) contains apostrophes in folders and filename', async () => {
    stubPostResponses();
    stubGetResponses();

    const unsafePath: string = '/Users/user/Projects/TEST\'FOLDER/TEST\'FILE.txt';

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: unsafePath,
        contentType: 'abc',
        debug: true
      }
    }));
    assert.strictEqual(loggerLogToStderrSpy.calledWith(`file name: TEST''FILE.txt...`), true);
  });

  it('should handle non existing content type', async () => {
    stubFs();
    stubPostResponses();
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'abc',
        debug: true
      }
    }), new CommandError('Specified content type \'abc\' doesn\'t exist on the target list'));
  });

  it('should handle list item update response error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Item update error." } } });

    stubFs();
    stubPostResponses(null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle list item field value update response error', async () => {
    const expectedResult: any = { "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "fsd", "HasException": false, "ItemId": 120 }, { "ErrorMessage": "check in comment x", "FieldName": "_CheckinComment", "FieldValue": "check in comment x", "HasException": true, "ItemId": 120 }] };

    stubFs();
    stubPostResponses(null, null, null, null, null, null, null, expectedResult);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        debug: true
      }
    }), new CommandError(`Update field value error: ${JSON.stringify(expectedResult.value)}`));
  });

  it('should handle file check in error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkin error." } } });

    stubFs();
    stubPostResponses(null, null, null, null, null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        checkInComment: 'abc',
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle approve list item response error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Approve error." } } });

    stubFs();
    stubPostResponses(null, null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true,
        verbose: true,
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should handle publish list item response error', async () => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Publish error." } } });

    stubFs();
    stubPostResponses(null, null, null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true,
        verbose: true,
        debug: true
      }
    }), new CommandError(expectedError));
  });

  it('should error when --publish used, but list moderation and minor ver enabled', async () => {
    const listSettingsResp = { "EnableMinorVersions": true, "EnableModeration": true, "EnableVersioning": true, "Id": "0c7dc8ec-5871-4ac9-962c-f856102b917b" };

    stubFs();
    stubPostResponses();
    stubGetResponses(null, null, null, null, listSettingsResp);

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        publish: true
      }
    }), new CommandError('The file cannot be published without approval. Moderation for this list is enabled. Use the --approve option instead of --publish to approve and publish the file'));
  });

  it('ignores global options when creating request data', async () => {
    stubFs();
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    await command.action(logger, {
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
    });
    assert.deepEqual(postRequests.secondCall.args[0].data, {
      bNewDocumentUpdate: true,
      checkInComment: '',
      formValues: [{ FieldName: 'Title', FieldValue: 'abc' }, { FieldName: 'ContentType', FieldValue: 'Picture' }]
    });
  });

  it('should perform single request upload for file up to 250 MB', async () => {
    stubFs();
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    sinonUtil.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 250 * 1024 * 1024 } as any); // 250 MB

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    });
    assert.notStrictEqual(postRequests.lastCall.args[0].url.indexOf(`/GetFolderByServerRelativeUrl('%2Fsites%2Fproject-x%2FShared%2520Documents%2Ft1')/Files/Add`), -1);
  });

  it('should perform chunk upload on files over 250 MB (debug)', async () => {
    stubFs();
    const postRequests: sinon.SinonStub = stubPostResponses();
    stubGetResponses();

    sinonUtil.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    });
    assert.notStrictEqual(postRequests.firstCall.args[0].url.indexOf('/StartUpload'), -1);
    assert.notStrictEqual(postRequests.getCalls()[2].args[0].url.indexOf('/ContinueUpload'), -1);
    assert.notStrictEqual(postRequests.lastCall.args[0].url.indexOf('/FinishUpload'), -1);
  });

  it('should cancel chunk upload on files over 250 MB on error', async () => {
    stubFs();
    stubGetResponses();
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/StartUpload') !== -1) {

          return { "d": { "StartUpload": "0" } };

        }
        else if ((opts.url as string).indexOf('/cancelupload') !== -1) {

          return { "d": { "CancelUpload": null } };

        }
        else if ((opts.url as string).indexOf('/ContinueUpload') !== -1) {

          throw { "error": "123" };

        }
      }
      throw 'Invalid request';
    });

    sinonUtil.restore([fs.statSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    } as any), new CommandError('123'));
  });

  it('should handle fail to read file error', async () => {
    stubGetResponses();
    stubPostResponses();

    sinonUtil.restore([fs.statSync, fs.openSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB
    sinon.stub(fs, 'openSync').throws(new Error('openSync error'));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    } as any), new CommandError('openSync error'));
  });

  it('should try closeSync on error', async () => {
    stubGetResponses();
    stubPostResponses();

    sinonUtil.restore([fs.statSync, fs.openSync, , fs.readSync, , fs.closeSync]);
    sinon.stub(fs, 'statSync').returns({ size: 251 * 1024 * 1024 } as any); // 250 MB
    sinon.stub(fs, 'openSync').returns(3);
    sinon.stub(fs, 'readSync').throws(new Error('readSync error'));
    sinon.stub(fs, 'closeSync').throws(new Error('failed to closeSync'));

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        debug: true,
        verbose: true
      }
    } as any), new CommandError('readSync error'));
  });

  it('should succeed updating list item metadata', async () => {
    stubFs();
    stubPostResponses();
    stubGetResponses();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        publish: true
      }
    });
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('sets field with the same name as a command option but different casing', async () => {
    stubFs();
    stubGetResponses();
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl(') > -1) {
        if ((opts.url as string).indexOf('/CheckOut') > -1) {
          return { "odata.null": true };
        }
        else if ((opts.url as string).indexOf('Add') > -1) {
          return { "CheckInComment": "", "CheckOutType": 0, "ContentTag": "{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428,159", "CustomizedPageStatus": 0, "ETag": "\"{B0BC16BB-C8D9-4A24-BC04-FB52045F8BEF},428\"", "Exists": true, "IrmEnabled": false, "Length": "165114", "Level": 255, "LinkingUri": null, "LinkingUrl": "", "MajorVersion": 51, "MinorVersion": 15, "Name": "MS365.jpg", "ServerRelativeUrl": "/sites/VelinDev/Shared Documents/t1/MS365.jpg", "TimeCreated": "2018-10-21T21:46:08Z", "TimeLastModified": "2018-10-25T23:49:52Z", "Title": "title4", "UIVersion": 26127, "UIVersionLabel": "51.15", "UniqueId": "b0bc16bb-c8d9-4a24-bc04-fb52045f8bef" };
        }
        else if ((opts.url as string).indexOf('ValidateUpdateListItem') > -1) {
          if (opts.data.formValues.filter((f: any) => f.FieldName === 'Folder').length > 0) {
            return { "value": [{ "ErrorMessage": null, "FieldName": "Title", "FieldValue": "title4", "HasException": false, "ItemId": 212 }] };
          }
          else {
            throw 'Field Folder missing';
          }
        }
        else if ((opts.url as string).indexOf('approve') > -1) {
          return { "odata.null": true };
        }
        else if ((opts.url as string).indexOf('publish') > -1) {
          return { "odata.null": true };
        }
        else if ((opts.url as string).indexOf('UndoCheckOut') > -1) {
          return { "odata.null": true };
        }
        else if ((opts.url as string).indexOf('CheckIn') > -1) {
          return { "odata.null": true };
        }
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        contentType: 'Picture',
        Title: 'abc',
        Folder: 'Folder',
        publish: true
      }
    } as any);
  });

  it('should succeed approve', async () => {
    stubFs();
    stubPostResponses();
    stubGetResponses();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        approve: true
      }
    });
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('should succeed when with checkout option', async () => {
    stubFs();
    stubPostResponses();
    stubGetResponses();

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }
    });
    assert.strictEqual(loggerLogSpy.notCalled, true);
  });

  it('should error if cannot rollback checkout (verbose)', async () => {
    stubFs();
    const expectedFileAddError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });

    stubPostResponses(null, expectedFileAddError, null, null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true,
        debug: true,
        verbose: true
      }
    }), new CommandError(expectedFileAddError));
  });

  it('should error if cannot rollback checkout', async () => {
    stubFs();
    const expectedFileAddError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File add error." } } });
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: Checkout Error." } } });

    stubPostResponses(null, expectedFileAddError, null, null, null, expectedError);
    stubGetResponses();

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        folder: 'Shared%20Documents/t1',
        path: 'C:\Users\Velin\Desktop\MS365.jpg',
        checkOut: true
      }
    }), new CommandError(expectedFileAddError));
  });

  it('fails validation if the webUrl option not valid url', async () => {
    const actual = await command.validate({ options: { webUrl: 'abc', folder: 'abc', path: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the wrong path option specified', async () => {
    sinon.stub(fs, 'existsSync').returns(false);
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if --approveComment specified, but not --approve', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc',
        approveComment: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if --publishComment specified, but not --publish', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc',
        publishComment: 'abc'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passed validation if options correct', async () => {
    sinon.stub(fs, 'existsSync').returns(true);
    const actual = await command.validate({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'abc',
        path: 'abc'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
