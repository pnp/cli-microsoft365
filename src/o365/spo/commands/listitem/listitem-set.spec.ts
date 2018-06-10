import commands from '../../commands';
import Command from '../../../../Command';
import { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./listitem-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.LISTITEM_SET, () => {

  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  const expectedTitle = `List Item 1`;

  const expectedId = 147;
  let actualId = 0;

  const expectedContentType = 'Item';
  let actualContentType = '';

  let postFakes = (opts: any) => {
    if (opts.url.indexOf('/common/oauth2/token') > -1) {
      return Promise.resolve('abc');
    }

    if (opts.url.indexOf('/_api/contextinfo') > -1) {
      return Promise.resolve({
        FormDigestValue: 'abc'
      });
    }
    if (opts.url.indexOf('ValidateUpdateListItem') > -1) {

      const bodyString = JSON.stringify(opts.body);
      const ctMatch = bodyString.match(/\"?FieldName\"?:\s*\"?ContentType\"?,\s*\"?FieldValue\"?:\s*\"?(\w*)\"?/i);
      actualContentType = ctMatch ? ctMatch[1] : "";
      if (bodyString.indexOf("fail updating me") > -1) return Promise.resolve({ value: [{ErrorMessage: 'failed updating'}] })
      return Promise.resolve({ value: [ { ItemId: expectedId }] });

    }
    if (opts.url.indexOf('_vti_bin/client.svc/ProcessQuery') > -1) {

      // requestObjectIdentity mock
      if (opts.body.indexOf('Name="Current"') > -1) {

        if (opts.url.indexOf('rejectme.com') > -1 ) {

          return Promise.reject('Failed request')

        }

        if (opts.url.indexOf('returnerror.com') > -1) {

          return Promise.resolve(JSON.stringify(
            [{"ErrorInfo": "error occurred"}]
          ))

        }

        return Promise.resolve(JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0",
              "LibraryVersion": "16.0.7618.1204",
              "ErrorInfo":null,
              "TraceCorrelationId": "3e3e629e-30cc-5000-9f31-cf83b8e70021"
            },
            {
              "_ObjectType_": "SP.Web", 
              "_ObjectIdentity_": "d704ae73-d5ed-459e-80b0-b8103c5fb6e0|8f2be65d-f195-4699-b0de-24aca3384ba9:site:0ead8b78-89e5-427f-b1bc-6e5a77ac191c:web:4c076c07-e3f1-49a8-ad01-dbb70b263cd7",
              "ServerRelativeUrl": "\\u002fsites\\u002fprojectx"
            }
          ])
        )

      }
      if (opts.body.indexOf('SystemUpdate') > -1) {

        if (opts.body.indexOf('systemUpdate error') > -1) {
          return Promise.resolve(
            'ErrorMessage": "systemUpdate error"}'
          )

        }

        actualId = expectedId;
        return Promise.resolve(
          ']SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.7618.1204","ErrorInfo":null,"TraceCorrelationId":"3e3e629e-f0e9-5000-9f31-c6758b453a4a"'
        )
      }
    }
    console.log('Invalid POST request')
    return Promise.reject('Invalid request');
  }

  let getFakes = (opts: any) => {
    if (opts.url.indexOf('contenttypes') > -1) {
      return Promise.resolve({ value: [ {Id: { StringValue: expectedContentType }, Name: "Item" } ] });
    }
    if (opts.url.indexOf('/items(') > -1) {
      actualId = opts.url.match(/\/items\((\d+)\)/i)[1];
      return Promise.resolve(
        {
          "Attachments": false,
          "AuthorId": 3,
          "ContentTypeId": "0x0100B21BD271A810EE488B570BE49963EA34",
          "Created": "2018-03-15T10:43:10Z",
          "EditorId": 3,
          "GUID": "ea093c7b-8ae6-4400-8b75-e2d01154dffc",
          "ID": actualId,
          "Modified": "2018-03-15T10:52:10Z",
          "Title": expectedTitle,
        }
      );
    }
    if (opts.url.indexOf('/id') > -1) {
      return Promise.resolve({ value: "f64041f2-9818-4b67-92ff-3bc5dbbef27e" });
    }
    console.log('Invalid GET request')
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
    assert.equal(command.name.startsWith(commands.LISTITEM_SET), true);
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
        assert.equal(telemetry.name, commands.LISTITEM_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to a SharePoint Online site first')));
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

  it('fails validation if listTitle and listId option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if listTitle and listId are specified together', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Demo List' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', listTitle: 'Demo List' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List' } });
    assert(actual);
  });

  it('fails validation if the listId option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', listId: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails to update a list item when \'fail me\' values are used', (done) => {

    actualId = 0;
    
    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      Title: "fail updating me"
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(actualId, 0);
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

  it('returns listItemInstance object when list item is updated with correct values', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    command.allowUnknownOptions();

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x', 
      Title: expectedTitle
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(actualId, expectedId);
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

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option \'Item\' is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      id: 47,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y', 
      contentType: 'Item',
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
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });

  });

  it('attempts to update the listitem with the contenttype of \'Item\' when content type option 0x01 is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      id: 47,
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
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });

  });

  it('fails to update the listitem when the specified contentType doesn\'t exist in the target list', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      id: 47,
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
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });

  });


  it('successfully updates the listitem when the systemUpdate parameter is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    actualId = 0;

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y', 
      Title: expectedTitle,
      systemUpdate: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert.equal(actualId, expectedId);
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

  it('fails to get _ObjecttIdentity_ when the systemUpdate parameter is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    actualId = 0;

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      id: 147,
      webUrl: 'https://rejectme.com/sites/project-y', 
      Title: expectedTitle,
      systemUpdate: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(actualId !== expectedId);
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

  it('fails to get _ObjecttIdentity_ when an error is returned by the _ObjectIdentity_ CSOM request and systemUpdate parameter is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    actualId = 0;

    let options: any = { 
      debug: false, 
      listTitle: 'Demo List', 
      id: 147,
      webUrl: 'https://returnerror.com/sites/project-y', 
      Title: expectedTitle,
      systemUpdate: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(actualId !== expectedId);
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

  it('fails to update the list item when systemUpdate parameter is specified', (done) => {

    sinon.stub(request, 'get').callsFake(getFakes);
    sinon.stub(request, 'post').callsFake(postFakes);

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    actualId = 0;

    let options: any = { 
      debug: true, 
      listTitle: 'Demo List', 
      id: 147,
      webUrl: 'https://contoso.sharepoint.com/sites/project-y', 
      Title: "systemUpdate error",
      contentType: "Item",
      systemUpdate: true
    }

    cmdInstance.action({ options: options }, () => {

      try {
        assert(actualId !== expectedId);
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
    assert(find.calledWith(commands.LISTITEM_SET));
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