import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./listitem-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_LIST, () => {

  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  const expectedArrayLength = 2;
  let returnArrayLength = 0;

  let postFakes = (opts: any) => {

    if (opts.url.indexOf('/common/oauth2/token') > -1) {
      returnArrayLength = 0;
      return Promise.resolve('abc');
    }
    if (opts.url.indexOf('_api/contextinfo') > -1) {
      returnArrayLength = 0;
      return Promise.resolve({
        FormDigestValue: 'abc'
      });
    }
    if (opts.url.indexOf('/GetItems') > -1) {
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
    if (opts.url.indexOf('/items') > -1) {
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
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.LISTITEM_LIST), true);
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
        assert.equal(telemetry.name, commands.LISTITEM_LIST);
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', title: 'Demo List' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
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

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('fails validation if title and id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if title and id are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', id: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Demo List' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', title: 'Demo List' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '935c13a0-cc53-4103-8b48-c1d0828eaa7f' } });
    assert(actual);
  });

  it('fails validation if query and fields are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><ViewFields><FieldRef Name="Title" /><FieldRef Name="Id" /></ViewFields></Query>', fields: 'Title,Id' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if query and pageSize are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><RowLimit>2</RowLimit></Query>', pageSize: 3 } });
    assert.notEqual(actual, true);
  });

  it('fails validation if query and pageNumber are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', query: '<Query><RowLimit>2</RowLimit></Query>', pageNumber: 3 } });
    assert.notEqual(actual, true);
  });

  it('fails validation if pageNumber is specified and pageSize is not', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageNumber: 3 } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specific pageSize is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageSize: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the specific pageNumber is not a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', title: 'Demo List', pageSize: 3, pageNumber: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('returns array of listItemInstance objects when a list of items is requested, and debug mode enabled', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a list of fields and a filter specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of json, and a pageNumber is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

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
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      fields: "Title,ID"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with an output type of text, and no fields specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      output: "text"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with no output type specified, and a list of fields specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      fields: "Title,ID",
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with a query specified, and output set to json, and debug mode is enabled', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: true, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      query: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>",
      output: "json"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
  });

  it('returns array of listItemInstance objects when a list of items is requested with a query specified, and debug mode is disabled', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      title: 'Demo List', 
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      query: "<View><Query><ViewFields><FieldRef Name='Title' /><FieldRef Name='Id' /></ViewFields><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>Demo List Item 1</Value></Eq></Where></Query></View>"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(returnArrayLength, expectedArrayLength);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
    
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
    assert(find.calledWith(commands.LISTITEM_LIST));
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
        id: 1,
        listId: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
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