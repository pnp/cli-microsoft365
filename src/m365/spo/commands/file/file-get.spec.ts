import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { PassThrough } from 'stream';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./file-get');

describe(commands.FILE_GET, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      fs.createWriteStream
    ]);
  });

  after(() => {
    sinonUtil.restore([
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

  it('excludes options from URL processing', () => {
    assert.deepStrictEqual((command as any).getExcludedOptionsWithUrls(), ['url']);
  });

  it('command correctly handles file get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        id: 'f09c4efe-b8c0-4e89-a166-03418661b89b'
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
        return Promise.resolve('Correct Url1');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
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
    const returnValue: string = 'BinaryFileString';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve(returnValue);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        asString: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(returnValue));
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

    command.action(logger, {
      options: {
        debug: true,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        asListItem: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith({
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
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    command.action(logger, {
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
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
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        url: '/sites/project-x/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
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
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
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
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        url: '/Documents/Test1.docx',
        webUrl: 'https://contoso.sharepoint.com'
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
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject(expectedError);
    });

    command.action(logger, {
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
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

  it('fails validation if path doesn\'t exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com/sites/project-x', id: 'b2307a39-e878-458b-bc90-03bc578531d6', asFile: true, path: 'abc' } }, commandInfo);
    sinonUtil.restore(fs.existsSync);
    assert.notStrictEqual(actual, true);
  });

  it('writeFile called when option --asFile is specified (verbose)', (done) => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('close');
    }, 5);

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve({
          data: responseStream
        });
      }

      return Promise.reject('Invalid request');
    });

    const options = {
      verbose: true,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      asFile: true,
      path: 'test1.docx',
      fileName: 'Test1.docx'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert(fsStub.calledOnce);
        assert.strictEqual(err, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          fs.createWriteStream
        ]);
      }
    });
  });

  it('fails when empty file is created file with --asFile is specified', (done) => {
    const mockResponse = `{"data": 123}`;
    const responseStream = new PassThrough();
    responseStream.write(mockResponse);
    responseStream.end(); //Mark that we pushed all the data.

    const writeStream = new PassThrough();
    const fsStub = sinon.stub(fs, 'createWriteStream').returns(writeStream as any);

    setTimeout(() => {
      writeStream.emit('error', "Writestream throws error");
    }, 5);

    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.resolve({
          data: responseStream
        });
      }

      return Promise.reject('Invalid request');
    });

    const options = {
      debug: false,
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      asFile: true,
      path: 'test1.docx',
      fileName: 'Test1.docx'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert(fsStub.calledOnce);
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Writestream throws error')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          fs.createWriteStream
        ]);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b' } }, commandInfo);
    assert(actual);
  });

  it('fails validation if the id or url option not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both id and url options are specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', url: '/sites/project-x/documents' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both path and fileName options are not specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asFile: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asFile and asListItem specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true, asListItem: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asFile and asString specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true, asString: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if asListItem and asString specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asListItem: true, asString: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if only asFile specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', path: 'abc', asFile: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only asListItem specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asListItem: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if only asString specified', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'f09c4efe-b8c0-4e89-a166-03418661b89b', asString: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
