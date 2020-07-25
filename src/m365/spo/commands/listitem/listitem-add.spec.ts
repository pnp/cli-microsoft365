import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { FolderExtensions } from '../../FolderExtensions';

describe(commands.LISTITEM_ADD, () => {
  let log: any[];
  let cmdInstance: any;
  let ensureFolderStub: sinon.SinonStub;

  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  let postFakes = (opts: any) => {
    if ((opts.url as string).indexOf('AddValidateUpdateItemUsingPath') > -1) {
      const bodyString = JSON.stringify(opts.body);
      const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail adding me") > -1) return Promise.resolve({ value: [] })
      return Promise.resolve({ value: [{ FieldName: "Id", FieldValue: expectedId }] });
    }

    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('contenttypes') > -1) {
      return Promise.resolve({ value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] });
    }
    if ((opts.url as string).indexOf('rootFolder') > -1) {
      return Promise.resolve({ ServerRelativeUrl: '/sites/project-xxx/Lists/Demo%20List' });
    }
    if ((opts.url as string).indexOf('/items(') > -1) {
      actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
      return Promise.resolve(
        {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "ID": actualId,
          "Modified": "2018-03-15T10:43:10Z",
          "Title": expectedTitle,
        }
      );
    }
    return Promise.reject('Invalid request');
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    ensureFolderStub = sinon.stub(FolderExtensions.prototype, 'ensureFolder').resolves();
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
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      FolderExtensions.prototype.ensureFolder,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listTitle: 'Demo List' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails to create a list item when \'fail me\' values are used', (done) => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: "fail adding me"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(actualId, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('returns listItemInstance object when list item is added with correct values', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    command.allowUnknownOptions();

    let options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: expectedTitle
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates list item in the list specified using ID', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    const options: any = {
      listId: 'cf8c72a1-0207-40ee-aebd-fca67d20bc8a',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      Title: expectedTitle
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(actualId, expectedId);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('attempts to create the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: true,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: expectedContentType,
      Title: expectedTitle
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(expectedContentType == actualContentType);
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('fails to create the listitem when the specified contentType doesn\'t exist in the target list', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = {
      debug: false,
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-y',
      contentType: "Unexpected content type",
      Title: expectedTitle
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(expectedContentType == actualContentType);
        done();
      }
      catch (e) {
        done(e);
      }
    });

  });

  it('should call ensure folder when folder arg specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({
      options: {
        debug: false,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2"
      }
    }, () => {
      try {
        assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
        assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should call ensure folder when folder arg specified (debug)', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3"
      }
    }, () => {
      try {
        assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
        assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should not have end \'/\' in the folder path when FolderPath.DecodedUrl ', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({
      options: {
        debug: true,
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      }
    }, () => {
      try {
        const addValidateUpdateItemUsingPathRequest = postStubs.getCall(postStubs.callCount - 1).args[0];
        const info = addValidateUpdateItemUsingPathRequest.body.listItemCreateInfo;
        assert.strictEqual(info.FolderPath.DecodedUrl, '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('ignores global options when creating request body', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(postFakes);

    cmdInstance.action({
      options: {
        debug: true,
        verbose: true,
        output: "text",
        listTitle: 'Demo List',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        Title: expectedTitle,
        contentType: expectedContentType,
        folder: "InsideFolder2/Folder3/"
      }
    }, () => {
      try {
        assert.deepEqual(postStubs.firstCall.args[0].body, {
          formValues: [{ FieldName: 'Title', FieldValue: 'List Item 1' }, { FieldName: 'ContentType', FieldValue: 'Item' }],
          listItemCreateInfo: { FolderPath: { DecodedUrl: '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3' } }
        });
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});