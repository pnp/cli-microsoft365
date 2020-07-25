import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./listitem-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_LIST, () => {
  let log: any[];
  let cmdInstance: any;

  const expectedArrayLength = 2;
  let returnArrayLength = 0;

  let postFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/GetItems') > -1) {
      returnArrayLength = 2;
      return Promise.resolve({value: 
        [{
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-08-15T13:43:12Z",
          "EditorId": 3,
          "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
          "ID": 1,
          "Modified": "2018-08-15T13:43:12Z",
          "Title": "Example item 1",
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
          "Title": "Example item 2",
        }]
      });
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if ((opts.url as string).indexOf('/items') > -1) {
      returnArrayLength = 2;
      return Promise.resolve({ value: 
        [{
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-08-15T13:43:12Z",
          "EditorId": 3,
          "GUID": "2b6bd9e0-3c43-4420-891e-20053e3c4664",
          "ID": 1,
          "Modified": "2018-08-15T13:43:12Z",
          "Title": "Example item 1",
        },
        {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-08-15T13:44:10Z",
          "EditorId": 3,
          "GUID": "47c5fc61-afb7-4081-aa32-f4386b8a86ea",
          "ID": 2,
          "Id": 2,
          "Modified": "2018-08-15T13:44:10Z",
          "Title": "Example item 2",
        }]
      });
    }
    returnArrayLength = 0;
    return Promise.reject('Invalid request');
  }
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub((command as any), 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc'
    }));
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_LIST), true);
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

  it('fails validation if title and id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', id: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', title: 'Demo List' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'foo' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert(actual);
  });

  it('fails validation if query and fields are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Id" /></ViewFields></Query>', fields: 'Title,Id' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if query and pageSize are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><RowLimit>2</RowLimit></Query>', pageSize: 3 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if query and pageNumber are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><RowLimit>2</RowLimit></Query>', pageNumber: 3 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if pageNumber is specified and pageSize is not', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageNumber: 3 } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specific pageSize is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageSize: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the specific pageNumber is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageSize: 3, pageNumber: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('returns array of listItemInstance objects when a list of items is requested, and debug mode enabled', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      output: "json",
      pageSize: 2,
      filter: "Title eq 'Demo list item",
      fields: "Title,ID"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      output: "json",
      pageSize: 2,
      pageNumber: 2,
      filter: "Title eq 'Demo list item",
      fields: "Title,ID"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a pageNumber is specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      output: "json",
      pageSize: 2,
      pageNumber: 2,
      fields: "Title,ID"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      fields: "Title,ID"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of text, and no fields specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      output: "text"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      fields: "Title,ID",
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with a query specified, and output set to json, and debug mode is enabled', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      query: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>",
      output: "json"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns array of listItemInstance objects when a list of items is requested with a query specified, and debug mode is disabled', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      query: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert.strictEqual(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));

    let options: any = { 
      debug: false, 
      id: '935c13a0-cc53-4103-8b48-c1d0828eaa7f', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      query: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    }

    cmdInstance.action({ options: options }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});