import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./file-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.FILE_GET, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.FILE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('command correctly handles file get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when output json option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('select123=') > -1) {
        return Promise.resolve('Correct Url1')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'b2307a39-e878-458b-bc90-03bc578531d6'
      }
    }, () => {
      try {
        assert('Correct Url');
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('retrieves file as binary string object', (done) => {
    let returnValue: string = 'BinaryFileString';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve(returnValue);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        asString: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(returnValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves and prints all details of file as ListItem object', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('?$expand=ListItemAllFields') > -1) {
        return Promise.resolve({
          "ListItemAllFields": {
            "FileSystemObjectType": 0,
            "Id": 4,
            "ServerRedirectedEmbedUri": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
            "ServerRedirectedEmbedUrl": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
            "ContentTypeId": "0x0101008E462E3ACE8DB844B3BEBF9473311889",
            "ComplianceAssetId": null,
            "Title": null,
            "ID": 4,
            "Created": "2018-02-05T09:42:36",
            "AuthorId": 1,
            "Modified": "2018-02-05T09:44:03",
            "EditorId": 1,
            "OData__CopySource": null,
            "CheckoutUserId": null,
            "OData__UIVersionString": "3.0",
            "GUID": "2054f49e-0f76-46d4-ac55-50e1c057941c"
          },
          "CheckInComment": "",
          "CheckOutType": 2,
          "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
          "CustomizedPageStatus": 0,
          "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
          "Exists": true,
          "IrmEnabled": false,
          "Length": "331673",
          "Level": 1,
          "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
          "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Documents/Test1.docx?d=wf09c4efeb8c04e89a16603418661b89b",
          "MajorVersion": 3,
          "MinorVersion": 0,
          "Name": "Opendag maart 2018.docx",
          "ServerRelativeUrl": "/sites/project-x/Documents/Test1.docx",
          "TimeCreated": "2018-02-05T08:42:36Z",
          "TimeLastModified": "2018-02-05T08:44:03Z",
          "Title": "",
          "UIVersion": 1536,
          "UIVersionLabel": "3.0",
          "UniqueId": "b2307a39-e878-458b-bc90-03bc578531d6"
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        asListItem: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "FileSystemObjectType": 0,
          "Id": 4,
          "ServerRedirectedEmbedUri": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
          "ServerRedirectedEmbedUrl": "https://contoso.sharepoint.com/sites/project-x/_layouts/15/WopiFrame.aspx?sourcedoc={b2307a39-e878-458b-bc90-03bc578531d6}&action=interactivepreview",
          "ContentTypeId": "0x0101008E462E3ACE8DB844B3BEBF9473311889",
          "ComplianceAssetId": null,
          "Title": null,
          "ID": 4,
          "Created": "2018-02-05T09:42:36",
          "AuthorId": 1,
          "Modified": "2018-02-05T09:44:03",
          "EditorId": 1,
          "OData__CopySource": null,
          "CheckoutUserId": null,
          "OData__UIVersionString": "3.0",
          "GUID": "2054f49e-0f76-46d4-ac55-50e1c057941c"
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when id option is passed', (done) => {
    const getStub: any = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, 'https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileById(\'0CD891EF-AFCE-4E55-B836-FCE03286CCCF\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when url option is passed', (done) => {
    const getStub: any = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?@f='%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx'`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when url option is passed to get file as list item', (done) => {
    const getStub: any = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        asListItem: true
      }
    }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, `https://contoso.sharepoint.com/sites/project-x/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?$expand=ListItemAllFields&@f='%2Fsites%2Fproject-x%2FDocuments%2FTest1.docx'`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses correct API url when tenant root URL option is passed', (done) => {
    const getStub: any = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileByServerRelativePath(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        url: '/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, () => {
      try {
        assert.strictEqual(getStub.lastCall.args[0].url, `https://contoso.sharepoint.com/_api/web/GetFileByServerRelativePath(DecodedUrl=@f)?@f='%2FDocuments%2FTest1.docx'`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle promise rejection', (done) => {
    const expectedError: any = JSON.stringify({ "odata.error": { "code": "-2130575338, Microsoft.SharePoint.SPException", "message": { "lang": "en-US", "value": "Error: File Not Found." } } });
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject(expectedError);
    });

    cmdInstance.action({
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      }
    }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err.message), JSON.stringify(expectedError));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if path doesn\'t exist', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com/sites/project-x', id: 'b2307a39-e878-458b-bc90-03bc578531d6', asFile: true, path: 'abc', fileName: 'test.docx' } });
    Utils.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('writeFile called when option --asFile is specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      asFile: true,
      path: '/Users/user/documents',
      fileName: 'Test1.docx'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(writeFileSyncStub.called)
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          fs.writeFileSync
        ]);
      }
    });
  });

  it('writeFile called when option --asFile is specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
    const options: Object = {
      debug: true,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      asFile: true,
      path: '/Users/user/documents',
      fileName: 'Test1.docx'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(writeFileSyncStub.called)
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          fs.writeFileSync
        ]);
      }
    });
  });

  it('writeFile not called when option --asFile and path is empty is specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const writeFileSyncStub = sinon.stub(fs, 'writeFileSync').callsFake(() => '');
    const options: Object = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      asFile: true,
      fileName: 'Test1.docx'
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(writeFileSyncStub.notCalled)
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          fs.writeFileSync
        ]);
      }
    });
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

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } });
    assert(actual);
  });

  it('fails validation if the id or url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', url: '/sites/project-x/documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both path and fileName options are not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asFile: true } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asFile and asListItem specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true, asListItem: true } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asFile and asString specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true, asString: true } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asListItem and asString specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asListItem: true, asString: true } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only asFile specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if only asListItem specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asListItem: true } });
    assert.strictEqual(actual, true);
  });

  it('passes validation if only asString specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asString: true } });
    assert.strictEqual(actual, true);
  });
});
