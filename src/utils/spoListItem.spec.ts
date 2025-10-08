import assert from 'assert';
import sinon from 'sinon';
import os from 'os';
import auth from '../Auth.js';
import { Logger } from '../cli/Logger.js';
import request from '../request.js';
import { sinonUtil } from '../utils/sinonUtil.js';
import { formatting } from './formatting.js';
import { ListItemAddOptions, ListItemListOptions, spoListItem } from './spoListItem.js';
import { urlUtil } from './urlUtil.js';
import { spo } from './spo.js';

describe('utils/spoListItem', () => {
  const webUrl = 'https://contoso.sharepoint.com/sites/project-x';
  const listUrl = 'sites/project-x/documents';
  const listServerRelativeUrl: string = urlUtil.getServerRelativePath(webUrl, listUrl);
  const listItemResponse = {
    value:
      [{
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-08-15T13:43:12Z",
        "EditorId": 3,
        "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
        "Id": 1,
        "ID": 1,
        "Modified": "2018-08-15T13:43:12Z",
        "Title": "Example item 1"
      },
      {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-08-15T13:44:10Z",
        "EditorId": 3,
        "GUID": "47c5fc61-afb7-4081-aa32-f4386b8a86ea",
        "Id": 2,
        "ID": 2,
        "Modified": "2018-08-15T13:44:10Z",
        "Title": "Example item 2"
      }]
  };

  let logger: Logger;
  let log: string[];
  let ensureFolderStub: sinon.SinonStub;

  const expectedArrayLength = 2;

  const listOperationPostFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('/GetItems') > -1) {
        return opts.data.query.ListItemCollectionPosition === undefined ? listItemResponse : { value: [] };
      }
    }
    throw 'Invalid request: ' + JSON.stringify(opts);;
  };

  const listOperationGetFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('/items') > -1 && (opts.url as string).indexOf('$top=6') > -1) {
        return { value: [] };
      }
      if ((opts.url as string).indexOf('/items') > -1) {
        return listItemResponse;
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items?$top=5000&$select=Title%2CID`) {
      return listItemResponse;
    }
    throw 'Invalid request: ' + JSON.stringify(opts);
  };


  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';
  const expectedTitle = `List Item 1`;

  const addOperationPostFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('AddValidateUpdateItemUsingPath') > -1) {
        const bodyString = JSON.stringify(opts.data);
        const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
        actualContentType = ctMatch ? ctMatch[1] : "";
        if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [{ ErrorMessage: 'failed updating', 'FieldName': 'Title', 'HasException': true }] }); }
        return { value: [{ FieldName: "Id", FieldValue: expectedId, HasException: false }] };
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/AddValidateUpdateItemUsingPath()`) {
      const bodyString = JSON.stringify(opts.data);
      const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail adding me") > -1) { return Promise.resolve({ value: [] }); }
      return { value: [{ FieldName: "Id", FieldValue: expectedId }] };
    }
    throw 'Invalid request';
  };

  const addOperationGetFakes = async (opts: any) => {
    if (opts.url.indexOf('/_api/web/lists') > -1) {
      if ((opts.url as string).indexOf('contenttypes') > -1) {
        return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
      }
      if ((opts.url as string).indexOf('rootFolder') > -1) {
        return { ServerRelativeUrl: '/sites/project-xxx/Lists/Demo%20List' };
      }
      if ((opts.url as string).indexOf('/items(') > -1) {
        actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
        return {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "Id": actualId,
          "ID": actualId,
          "Modified": "2018-03-15T10:43:10Z",
          "Title": expectedTitle
        };
      }
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/contenttypes?$select=Name,Id`) {
      return { value: [{ Id: { StringValue: expectedContentType }, Name: "Item" }] };
    }
    if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/GetList('${formatting.encodeQueryParameter(listServerRelativeUrl)}')/items(147)`) {
      actualId = parseInt(opts.url.match(/\/items\((\d+)\)/i)[1]);
      return {
        "Attachments": false,
        "AuthorId": 3,
        "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
        "Created": "2018-03-15T10:43:10Z",
        "EditorId": 3,
        "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
        "Id": actualId,
        "ID": actualId,
        "Modified": "2018-03-15T10:43:10Z",
        "Title": expectedTitle
      };
    }
    throw 'Invalid request';
  };

  before(() => {
    auth.connection.active = true;
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    ensureFolderStub = sinon.stub(spo, 'ensureFolder').resolves();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post
    ]);
    auth.connection.spoUrl = undefined;
    auth.connection.spoTenantId = undefined;
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('returns array of listItemInstance objects when a list of items is requested, and debug mode enabled', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x'
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested, and a list of fields and a filter specified', async () => {
    const listTitle = `Test'list`;
    const filter = `Title eq 'Demo list item'`;
    const fields = 'Title,ID';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/project-x/_api/web/lists/getByTitle('${formatting.encodeQueryParameter(listTitle)}')/items?$top=2&$filter=${encodeURIComponent(filter)}&$select=${formatting.encodeQueryParameter(fields)}`) {
        return listItemResponse;
      }
      throw 'Invalid request';
    });

    const options: ListItemListOptions = {
      listTitle: listTitle,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      pageSize: 2,
      filter: filter,
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, a page number specified, a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      pageSize: 2,
      pageNumber: 2,
      filter: "Title eq 'Demo list item",
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns empty array of listItemInstance objects when a list of items is requested with an output type of json, a page number specified, a list of fields and a filter specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      pageSize: 3,
      pageNumber: 2,
      filter: "Title eq 'Demo list item",
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, 0);
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a pageNumber is specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      pageSize: 2,
      pageNumber: 2,
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items by list url is requested with no output type specified, and a list of fields specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listUrl: listUrl,
      webUrl: webUrl,
      fields: ['Title', 'ID']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, a list of fields with lookup field specified', async () => {
    sinon.stub(request, 'get').callsFake(opts => {
      if ((opts.url as string).indexOf('$expand=') > -1) {
        return Promise.resolve({
          value:
            [{
              "ID": 1,
              "Modified": "2018-08-15T13:43:12Z",
              "Title": "Example item 1",
              "Company": { "Title": "Contoso" }
            },
            {
              "ID": 2,
              "Modified": "2018-08-15T13:44:10Z",
              "Title": "Example item 2",
              "Company": { "Title": "Fabrikam" }
            }]
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      fields: ['Title', 'Modified', 'Company/Title']
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.deepStrictEqual(JSON.stringify(listItems), JSON.stringify([
      {
        "Modified": "2018-08-15T13:43:12Z",
        "Title": "Example item 1",
        "Company": { "Title": "Contoso" }
      },
      {
        "Modified": "2018-08-15T13:44:10Z",
        "Title": "Example item 2",
        "Company": { "Title": "Fabrikam" }
      }
    ]));
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and no fields specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x'
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified, and output set to json, and debug mode is enabled', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('returns array of listItemInstance objects when a list of items is requested with a camlQuery specified', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(listOperationPostFakes);

    const options: ListItemListOptions = {
      listTitle: 'Demo List',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    };

    const listItems = await spoListItem.getListItems(options, logger, true);
    assert.strictEqual(listItems.length, expectedArrayLength);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(listOperationGetFakes);
    sinon.stub(request, 'post').callsFake(() => Promise.reject(new Error('An error has occurred')));

    const options: ListItemListOptions = {
      listId: '935c13a0-cc53-4103-8b48-c1d0828eaa7f',
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      camlQuery: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    };

    await assert.rejects(spoListItem.getListItems(options, logger, true), new Error('An error has occurred'));
  });

  it('fails to create a list item when \'fail me\' values are used', async () => {
    actualId = 0;

    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      fieldValues: { Title: "fail adding me" }
    };

    await assert.rejects(spoListItem.addListItem(options, logger, true, true), new Error(`Creating the item failed with the following errors: ${os.EOL}- Title - failed updating`));
    assert.strictEqual(actualId, 0);
  });

  it('returns listItemInstance object when list item is added with correct values', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle }
    };

    await spoListItem.addListItem(options, logger, true, true);
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using ID', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listId: 'cf8c72a1-0207-40ee-aebd-fca67d20bc8a',
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle }
    };

    await spoListItem.addListItem(options, logger, true, true);
    assert.strictEqual(actualId, expectedId);
  });

  it('creates list item in the list specified using URL', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listUrl: listUrl,
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle }
    };

    await spoListItem.addListItem(options, logger, true, true);
    assert.strictEqual(actualId, expectedId);
  });


  it('attempts to create the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: expectedContentType,
      fieldValues: { Title: expectedTitle }
    };

    await spoListItem.addListItem(options, logger, true, true);
    assert(expectedContentType === actualContentType);
  });

  it('fails to create the listitem when the specified contentType doesn\'t exist in the target list', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      contentType: "Unexpected content type",
      fieldValues: { Title: expectedTitle }
    };

    await assert.rejects(spoListItem.addListItem(options, logger, true, true), new Error("Specified content type 'Unexpected content type' doesn't exist on the target list"));
  });

  it('should call ensure folder when folder arg specified', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle },
      contentType: expectedContentType,
      folder: "InsideFolder2"
    };

    await spoListItem.addListItem(options, logger, true, true);

    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2');
  });

  it('should call ensure folder when folder arg specified (debug)', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle },
      contentType: expectedContentType,
      folder: "InsideFolder2/Folder3"
    };

    await spoListItem.addListItem(options, logger, true, true);

    assert.strictEqual(ensureFolderStub.lastCall.args[0], 'https://contoso.sharepoint.com/sites/project-x');
    assert.strictEqual(ensureFolderStub.lastCall.args[1], '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });

  it('should not have end \'/\' in the folder path when FolderPath.DecodedUrl ', async () => {
    sinon.stub(request, 'get').callsFake(addOperationGetFakes);
    const postStubs = sinon.stub(request, 'post').callsFake(addOperationPostFakes);

    const options: ListItemAddOptions = {
      listTitle: 'Demo List',
      webUrl: webUrl,
      fieldValues: { Title: expectedTitle },
      contentType: expectedContentType,
      folder: "InsideFolder2/Folder3/"
    };

    await spoListItem.addListItem(options, logger, true, true);

    const addValidateUpdateItemUsingPathRequest = postStubs.getCall(postStubs.callCount - 1).args[0];
    const info = addValidateUpdateItemUsingPathRequest.data.listItemCreateInfo;
    assert.strictEqual(info.FolderPath.DecodedUrl, '/sites/project-xxx/Lists/Demo%20List/InsideFolder2/Folder3');
  });
});