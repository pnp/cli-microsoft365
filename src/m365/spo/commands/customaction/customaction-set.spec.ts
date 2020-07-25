import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./customaction-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.CUSTOMACTION_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let defaultCommandOptions: any;
  let initDefaultPostStubs = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });
  }

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
    defaultCommandOptions = {
      url: 'https://contoso.sharepoint.com',
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      title: 'title'
    }
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      (command as any).updateCustomAction,
      (command as any).searchAllScopes
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
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has a permissionsKindMap', () => {
    assert.strictEqual((command as any)['permissionsKindMap'].length, 37);
  });

  it('correct https body send when custom action with location StandardMenu', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 102,
      location: 'Microsoft.SharePoint.StandardMenu',
      description: 'description1',
      group: 'SiteActions',
      actionUrl: '~site/Shared%20Documents/Forms/AllItems.aspx'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Group: 'SiteActions',
            Description: 'description1',
            Sequence: 102,
            Url: '~site/Shared%20Documents/Forms/AllItems.aspx'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location ClientSideExtension.ApplicationCustomizer', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 101,
      location: 'ClientSideExtension.ApplicationCustomizer',
      description: 'description1',
      clientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
      clientSideComponentProperties: '{"testMessage":"Test message"}',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'ClientSideExtension.ApplicationCustomizer',
            Description: 'description1',
            Sequence: 101,
            ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
            ClientSideComponentProperties: '{"testMessage":"Test message"}'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location ClientSideExtension.ListViewCommandSet', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 103,
      location: 'ClientSideExtension.ListViewCommandSet',
      description: 'description1',
      clientSideComponentId: 'db3e6e35-363c-42b9-a254-ca661e437848',
      clientSideComponentProperties: '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}',
      registrationId: 100,
      registrationType: 'List'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'ClientSideExtension.ListViewCommandSet',
            Description: 'description1',
            Sequence: 103,
            ClientSideComponentId: 'db3e6e35-363c-42b9-a254-ca661e437848',
            ClientSideComponentProperties: '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}',
            RegistrationId: '100',
            RegistrationType: 1
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location EditControlBlock', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 104,
      location: 'EditControlBlock',
      description: 'description1',
      actionUrl: 'javascript:(function(){ return console.log("CLI for Microsoft 365 rocks!"); })();',
      registrationId: 101,
      registrationType: 'List'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'EditControlBlock',
            Description: 'description1',
            Sequence: 104,
            Url: 'javascript:(function(){ return console.log("CLI for Microsoft 365 rocks!"); })();',
            RegistrationId: '101',
            RegistrationType: 1
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location ScriptLink', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 105,
      location: 'ScriptLink',
      description: 'description1',
      scriptSrc: '~sitecollection/SiteAssets/YourScript.js',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'ScriptLink',
            Description: 'description1',
            Sequence: 105,
            ScriptSrc: '~sitecollection/SiteAssets/YourScript.js'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location ScriptLink and ScriptBlock', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 108,
      location: 'ScriptLink',
      scriptBlock: '(function(){ return console.log("Hello CLI for Microsoft 365!"); })();'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'ScriptLink',
            Sequence: 108,
            ScriptBlock: '(function(){ return console.log("Hello CLI for Microsoft 365!"); })();'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with location CommandUI.Ribbon', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 106,
      location: 'CommandUI.Ribbon',
      imageUrl: '_layouts/15/images/placeholder32x32.png',
      description: 'description1',
      commandUIExtension: '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location = "Ribbon.List.Share.Controls._children"><Button Id = "Ribbon.List.Share.GetItemsCountButton" Alt = "Get list items count" Sequence = "11" Command = "Invoke_GetItemsCountButtonRequest" LabelText = "Get Items Count" TemplateAlias = "o1" Image32by32 = "_layouts/15/images/placeholder32x32.png" Image16by16 = "_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command = "Invoke_GetItemsCountButtonRequest" CommandAction = "javascript: alert(ctx.TotalListItems);" EnabledScript = "javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'CommandUI.Ribbon',
            Description: 'description1',
            Sequence: 106,
            ImageUrl: '_layouts/15/images/placeholder32x32.png',
            CommandUIExtension: '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location = "Ribbon.List.Share.Controls._children"><Button Id = "Ribbon.List.Share.GetItemsCountButton" Alt = "Get list items count" Sequence = "11" Command = "Invoke_GetItemsCountButtonRequest" LabelText = "Get Items Count" TemplateAlias = "o1" Image32by32 = "_layouts/15/images/placeholder32x32.png" Image16by16 = "_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command = "Invoke_GetItemsCountButtonRequest" CommandAction = "javascript: alert(ctx.TotalListItems);" EnabledScript = "javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https body send when custom action with delegated rights', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 107,
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      actionUrl: '~site/SitePages/Home.aspx',
      rights: 'AddListItems,DeleteListItems,ManageLists',
      debug: true
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          body: {
            Title: 'title1',
            Name: 'name1',
            Location: 'Microsoft.SharePoint.StandardMenu',
            Group: 'SiteActions',
            Url: '~site/SitePages/Home.aspx',
            Sequence: 107,
            Rights: { High: '0', Low: '2058' }
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves and prints the added user custom actions details when verbose specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    defaultCommandOptions.verbose = true;

    cmdInstance.action({ options: defaultCommandOptions }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(sinon.match('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updateCustomAction called once when scope is Web', (done) => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake((opts) => {  
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledOnce, 'postRequestSpy.calledOnce');
        assert(updateCustomActionSpy.calledWith({
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          url: 'https://contoso.sharepoint.com',
          scope: 'Web'
        }), 'updateCustomActionSpy.calledWith');
        assert(updateCustomActionSpy.calledOnce, 'updateCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updateCustomAction called once when scope is Site', (done) => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: 'Site'
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(postRequestSpy.calledOnce, 'postRequestSpy.calledOnce');
        assert(updateCustomActionSpy.calledWith(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site'
          }), 'updateCustomActionSpy.calledWith');
        assert(updateCustomActionSpy.calledOnce, 'updateCustomActionSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updateCustomAction called once when scope is All, but item found on web level', (done) => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake((opts) => {  
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');

    cmdInstance.action({
      options: {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com',
        scope: 'All',
        title: 'title'
      }
    }, () => {
      try {
        assert(postRequestSpy.calledOnce, 'postRequest');
        assert(updateCustomActionSpy.calledOnce, 'updateCustomAction');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updateCustomAction called twice when scope is All, but item not found on web level', (done) => {
    let postRequestSpy = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');

    cmdInstance.action({
      options: {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        url: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(postRequestSpy.calledTwice, 'postRequest');
        assert(updateCustomActionSpy.calledTwice, 'updateCustomAction');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('searchAllScopes called when scope is All', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: "All"
    }

    cmdInstance.action({ options: options }, () => {
      try {
        assert(searchAllScopesSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com'
          })), 'searchAllScopesSpy.calledWith');
        assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: defaultCommandOptions
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('searchAllScopes correctly handles custom action Web odata.null when All scope specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options: Object = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      url: 'https://contoso.sharepoint.com',
      scope: "All",
      debug: true
    }

    cmdInstance.action({
      options: options
    }, () => {
      try {
        assert(searchAllScopesSpy.calledOnce);
        assert(updateCustomActionSpy.calledTwice);
        assert(updateCustomActionSpy.calledWith(sinon.match(
          {
            id: 'b2307a39-e878-458b-bc90-03bc578531d6',
            url: 'https://contoso.sharepoint.com',
            scope: 'Site'
          })), 'searchAllScopesSpy.calledWith');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(`Custom action with id ${actionId} not found`));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('searchAllScopes correctly handles web custom action reject request', (done) => {
    const err = 'Invalid custom action request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }
      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('searchAllScopes correctly handles site custom action reject request', (done) => {
    const err = 'Invalid custom action request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    cmdInstance.action({
      options: {
        debug: false,
        verbose: true,
        id: actionId,
        url: 'https://contoso.sharepoint.com',
        scope: 'All'
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

  it('offers autocomplete for the registrationType option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--registrationType') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the rights option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--rights') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the scope option', () => {
    const options = (command.options() as CommandOption[]);
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--scope') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('getRegistrationType returns 0 if registrationType wrong value', () => {
    let registrationType: number = (command as any)['getRegistrationType']('abc');
    assert(registrationType === 0);
  });

  it('getRegistrationType returns 1 if registrationType value is List', () => {
    let registrationType: number = (command as any)['getRegistrationType']('List');
    assert(registrationType === 1);
  });

  it('getRegistrationType returns 2 if registrationType value is ContentType', () => {
    let registrationType: number = (command as any)['getRegistrationType']('ContentType');
    assert(registrationType === 2);
  });

  it('getRegistrationType returns 3 if registrationType value is ProgId', () => {
    let registrationType: number = (command as any)['getRegistrationType']('ProgId');
    assert(registrationType === 3);
  });

  it('getRegistrationType returns 4 if registrationType value is FileType', () => {
    let registrationType: number = (command as any)['getRegistrationType']('FileType');
    assert(registrationType === 4);
  });

  it('fails validation if no other fields specified than url, id, scope', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '058140e3-0e37-44fc-a1d3-79c487d371a3', url:'https://contoso.sharepoint.com'} });
    assert.strictEqual(actual, 'Please specify option to be updated');
  });

  it('fails if the specified URL is invalid', () => {
    defaultCommandOptions.url = 'foo';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.notStrictEqual(actual, true);
  });

  it('getRegistrationType returns 1 if registrationType value is List', () => {
    let registrationType: number = (command as any)['getRegistrationType']('List');
    assert(registrationType === 1);
  });

  it('should map independently location', () => {
    let result: number = (command as any)['mapRequestBody']({location: 'abc'});
    assert(JSON.stringify(result) === `{"Location":"abc"}`);
  });

  it('should map independently name', () => {
    let result: number = (command as any)['mapRequestBody']({name: 'abc'});
    assert(JSON.stringify(result) === `{"Name":"abc"}`);
  });

  it('should map independently title', () => {
    let result: number = (command as any)['mapRequestBody']({title: 'abc'});
    assert(JSON.stringify(result) === `{"Title":"abc"}`);
  });

  it('should map independently group', () => {
    let result: number = (command as any)['mapRequestBody']({group: 'abc'});
    assert(JSON.stringify(result) === `{"Group":"abc"}`);
  });

  it('fails validation if invalid id', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '1', url:'https://contoso.sharepoint.com'} });
    assert.strictEqual(actual, '1 is not valid. Custom action id (Guid) expected');
  });

  it('fails validation if undefined id', () => {
    const actual = (command.validate() as CommandValidate)({ options: { url:'https://contoso.sharepoint.com'} });
    assert.strictEqual(actual, 'undefined is not valid. Custom action id (Guid) expected');
  });

  it('fails if non existing PermissionKind rights specified', () => {
    defaultCommandOptions.rights = 'abc';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.strictEqual(actual, `Rights option '${defaultCommandOptions.rights}' is not recognized as valid PermissionKind choice. Please note it is case-sensitive`);
  });

  it('has correct PermissionKind rights specified', () => {
    defaultCommandOptions.rights = 'FullMask';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert(actual === true);
  });

  it('fails if clientSideComponentId is not a valid GUID', () => {
    defaultCommandOptions.clientSideComponentId = 'abc';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.strictEqual(actual, `ClientSideComponentId ${defaultCommandOptions.clientSideComponentId} is not a valid GUID`);
  });

  it('fails if the sequence value less than 0', () => {
    defaultCommandOptions.sequence = -1;
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.strictEqual(actual, `Invalid option sequence. Expected value in range from 0 to 65536`);
  });

  it('fails if the sequence value is higher than 65536', () => {
    defaultCommandOptions.sequence = 65537;
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.strictEqual(actual, `Invalid option sequence. Expected value in range from 0 to 65536`);
  });

  it('fails if both option scriptSrc and scriptBlock specified', () => {
    defaultCommandOptions.location = 'ScriptLink';
    defaultCommandOptions.scriptSrc = 'abc';
    defaultCommandOptions.scriptBlock = 'abc';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert.strictEqual(actual, `Either option scriptSrc or scriptBlock can be specified, but not both`);
  });

  it('success if location that requires group option is set, but group is also set', () => {
    defaultCommandOptions.location = 'Microsoft.SharePoint.StandardMenu';
    defaultCommandOptions.group = 'SiteActions';
    const actual = (command.validate() as CommandValidate)({ options: defaultCommandOptions });
    assert(actual === true);
  });

  it('calls the correct endpoint (url) when scope is Web', (done) => {
    const postRequestSpy = initDefaultPostStubs();

    cmdInstance.action({ options: defaultCommandOptions }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          url: `https://contoso.sharepoint.com/_api/Web/UserCustomActions('${defaultCommandOptions.id}')`,
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the correct endpoint (url) when scope is Site', (done) => {
    const postRequestSpy = initDefaultPostStubs();

    defaultCommandOptions.scope = "Site";
    cmdInstance.action({ options: defaultCommandOptions }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          url: `https://contoso.sharepoint.com/_api/Site/UserCustomActions('${defaultCommandOptions.id}')`,
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should have X-HTTP-Method MERGE header when scope is Web', (done) => {
    const postRequestSpy = initDefaultPostStubs();

    cmdInstance.action({ options: defaultCommandOptions }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          headers: {
            accept: 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should have X-HTTP-Method MERGE when scope is Site', (done) => {
    const postRequestSpy = initDefaultPostStubs();

    defaultCommandOptions.scope = "Site";
    cmdInstance.action({ options: defaultCommandOptions }, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          headers: {
            accept: 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE'
          }
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects invalid string scope', () => {
    defaultCommandOptions.scope = 'All1';
    const actual = (command.validate() as CommandValidate)({
      options: defaultCommandOptions
    });
    assert.strictEqual(actual, `${defaultCommandOptions.scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options: defaultCommandOptions
      });
    assert(actual === true);
  });

  it('supports specifying scope', () => {
    const options = (command.options() as CommandOption[]);
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('doesn\'t fail if the parent doesn\'t define options', () => {
    sinon.stub(Command.prototype, 'options').callsFake(() => { return []; });
    const options = (command.options() as CommandOption[]);
    Utils.restore(Command.prototype.options);
    assert(options.length > 0);
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
});