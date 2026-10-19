import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './customaction-set.js';

describe(commands.CUSTOMACTION_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let defaultCommandOptions: any;
  const initDefaultPostStubs = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    defaultCommandOptions = {
      webUrl: 'https://contoso.sharepoint.com',
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      title: 'title'
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      (command as any).updateCustomAction,
      (command as any).searchAllScopes,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CUSTOMACTION_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has a permissionsKindMap', () => {
    assert.strictEqual((command as any)['permissionsKindMap'].length, 37);
  });

  it('correct https body send when custom action with location StandardMenu', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 102,
      location: 'Microsoft.SharePoint.StandardMenu',
      description: 'description1',
      group: 'SiteActions',
      actionUrl: '~site/Shared%20Documents/Forms/AllItems.aspx'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Group: 'SiteActions',
        Description: 'description1',
        Sequence: 102,
        Url: '~site/Shared%20Documents/Forms/AllItems.aspx'
      }
    })));
  });

  it('correct https data send when custom action with location ClientSideExtension.ApplicationCustomizer', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 101,
      location: 'ClientSideExtension.ApplicationCustomizer',
      description: 'description1',
      clientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
      clientSideComponentProperties: '{"testMessage":"Test message"}',
      debug: true
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'ClientSideExtension.ApplicationCustomizer',
        Description: 'description1',
        Sequence: 101,
        ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
        ClientSideComponentProperties: '{"testMessage":"Test message"}'
      }
    })));
  });

  it('correct https data send when custom action with location ClientSideExtension.ListViewCommandSet', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 103,
      location: 'ClientSideExtension.ListViewCommandSet',
      description: 'description1',
      clientSideComponentId: 'db3e6e35-363c-42b9-a254-ca661e437848',
      clientSideComponentProperties: '{"sampleTextOne":"One item is selected in the list.", "sampleTextTwo":"This command is always visible."}',
      registrationId: '100',
      registrationType: 'List'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
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
  });

  it('correct https data send when custom action with location EditControlBlock', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 104,
      location: 'EditControlBlock',
      description: 'description1',
      actionUrl: 'javascript:(function(){ return console.log("CLI for Microsoft 365 rocks!"); })();',
      registrationId: '101',
      registrationType: 'List'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
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
  });

  it('correct https data send when custom action with location ScriptLink', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 105,
      location: 'ScriptLink',
      description: 'description1',
      scriptSrc: '~sitecollection/SiteAssets/YourScript.js',
      scope: 'Site'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'ScriptLink',
        Description: 'description1',
        Sequence: 105,
        ScriptSrc: '~sitecollection/SiteAssets/YourScript.js'
      }
    })));
  });

  it('correct https data send when custom action with location ScriptLink and ScriptBlock', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 108,
      location: 'ScriptLink',
      scriptBlock: '(function(){ return console.log("Hello CLI for Microsoft 365!"); })();'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'ScriptLink',
        Sequence: 108,
        ScriptBlock: '(function(){ return console.log("Hello CLI for Microsoft 365!"); })();'
      }
    })));
  });

  it('correct https data send when custom action with location CommandUI.Ribbon', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 106,
      location: 'CommandUI.Ribbon',
      imageUrl: '_layouts/15/images/placeholder32x32.png',
      description: 'description1',
      commandUIExtension: '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location = "Ribbon.List.Share.Controls._children"><Button Id = "Ribbon.List.Share.GetItemsCountButton" Alt = "Get list items count" Sequence = "11" Command = "Invoke_GetItemsCountButtonRequest" LabelText = "Get Items Count" TemplateAlias = "o1" Image32by32 = "_layouts/15/images/placeholder32x32.png" Image16by16 = "_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command = "Invoke_GetItemsCountButtonRequest" CommandAction = "javascript: alert(ctx.TotalListItems);" EnabledScript = "javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>',
      scope: 'Web'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'CommandUI.Ribbon',
        Description: 'description1',
        Sequence: 106,
        ImageUrl: '_layouts/15/images/placeholder32x32.png',
        CommandUIExtension: '<CommandUIExtension><CommandUIDefinitions><CommandUIDefinition Location = "Ribbon.List.Share.Controls._children"><Button Id = "Ribbon.List.Share.GetItemsCountButton" Alt = "Get list items count" Sequence = "11" Command = "Invoke_GetItemsCountButtonRequest" LabelText = "Get Items Count" TemplateAlias = "o1" Image32by32 = "_layouts/15/images/placeholder32x32.png" Image16by16 = "_layouts/15/images/placeholder16x16.png" /></CommandUIDefinition></CommandUIDefinitions><CommandUIHandlers><CommandUIHandler Command = "Invoke_GetItemsCountButtonRequest" CommandAction = "javascript: alert(ctx.TotalListItems);" EnabledScript = "javascript: function checkEnable() { return (true);} checkEnable();"/></CommandUIHandlers></CommandUIExtension>'
      }
    })));
  });

  it('correct https data send when custom action with delegated rights', async () => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 107,
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      actionUrl: '~site/SitePages/Home.aspx',
      rights: 'AddListItems,DeleteListItems,ManageLists',
      debug: true
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledWith(sinon.match({
      data: {
        Title: 'title1',
        Name: 'name1',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Group: 'SiteActions',
        Url: '~site/SitePages/Home.aspx',
        Sequence: 107,
        Rights: { High: '0', Low: '2058' }
      }
    })));
  });

  it('updateCustomAction called once when scope is Web', async () => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      title: 'title1',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledOnce, 'postRequestSpy.calledOnce');
    assert(updateCustomActionSpy.calledWithMatch({
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      title: 'title1',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    }), 'updateCustomActionSpy.calledWithMatch');
    assert(updateCustomActionSpy.calledOnce, 'updateCustomActionSpy.calledOnce');
  });

  it('updateCustomAction called once when scope is Site', async () => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      title: 'title1',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(postRequestSpy.calledOnce, 'postRequestSpy.calledOnce');
    assert(updateCustomActionSpy.calledWithMatch(
      {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }), 'updateCustomActionSpy.calledWithMatch');
    assert(updateCustomActionSpy.calledOnce, 'updateCustomActionSpy.calledOnce');
  });

  it('updateCustomAction called once when scope is All, but item found on web level', async () => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        title: 'title'
      })
    });
    assert(postRequestSpy.calledOnce, 'postRequest');
    assert(updateCustomActionSpy.calledOnce, 'updateCustomAction');
  });

  it('updateCustomAction called twice when scope is All, but item not found on web level', async () => {
    const postRequestSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com'
      })
    });
    assert(postRequestSpy.calledTwice, 'postRequest');
    assert(updateCustomActionSpy.calledTwice, 'updateCustomAction');
  });

  it('searchAllScopes called when scope is All', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      title: 'title1',
      webUrl: 'https://contoso.sharepoint.com',
      scope: "All"
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(searchAllScopesSpy.calledWith(sinon.match(
      {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com'
      })), 'searchAllScopesSpy.calledWith');
    assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse(defaultCommandOptions)
    });
    assert(loggerLogSpy.notCalled);
  });

  it('searchAllScopes correctly handles custom action Web odata.null when All scope specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return 'abc';
      }

      throw 'Invalid request';
    });

    const updateCustomActionSpy = sinon.spy((command as any), 'updateCustomAction');
    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      title: 'title1',
      webUrl: 'https://contoso.sharepoint.com',
      scope: "All",
      debug: true
    };

    await command.action(logger, {
      options: commandOptionsSchema.parse(options)
    });
    assert(searchAllScopesSpy.calledOnce);
    assert(updateCustomActionSpy.calledTwice);
    assert(updateCustomActionSpy.calledWith(sinon.match(
      {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      })), 'searchAllScopesSpy.calledWith');
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        id: actionId,
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      })
    });
    assert(loggerLogToStderrSpy.calledWith(`Custom action with id ${actionId} not found`));
  });

  it('searchAllScopes correctly handles web custom action reject request', async () => {
    const err = 'Invalid custom action request';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        throw err;
      }
      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    try {
      await assert.rejects(command.action(logger, {
        options: commandOptionsSchema.parse({
          id: actionId,
          title: 'title1',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'All'
        })
      }), new CommandError(err));
    }
    finally {
      sinonUtil.restore(request.post);
    }
  });

  it('searchAllScopes correctly handles site custom action reject request', async () => {
    const err = 'Invalid custom action request';
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        id: actionId,
        title: 'title1',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      })
    }), new CommandError(err));
  });

  it('getRegistrationType returns 0 if registrationType wrong value', () => {
    const registrationType: number = (command as any)['getRegistrationType']('abc');
    assert(registrationType === 0);
  });

  it('getRegistrationType returns 1 if registrationType value is List', () => {
    const registrationType: number = (command as any)['getRegistrationType']('List');
    assert(registrationType === 1);
  });

  it('getRegistrationType returns 2 if registrationType value is ContentType', () => {
    const registrationType: number = (command as any)['getRegistrationType']('ContentType');
    assert(registrationType === 2);
  });

  it('getRegistrationType returns 3 if registrationType value is ProgId', () => {
    const registrationType: number = (command as any)['getRegistrationType']('ProgId');
    assert(registrationType === 3);
  });

  it('getRegistrationType returns 4 if registrationType value is FileType', () => {
    const registrationType: number = (command as any)['getRegistrationType']('FileType');
    assert(registrationType === 4);
  });

  it('fails validation if no other fields specified than url, id, scope', () => {
    const actual = commandOptionsSchema.safeParse({ id: '058140e3-0e37-44fc-a1d3-79c487d371a3', webUrl: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, false);
  });

  it('fails if the specified URL is invalid', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'foo',
      title: 'title'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('should map independently location', () => {
    const result: number = (command as any)['mapRequestBody']({ location: 'abc' });
    assert(JSON.stringify(result) === `{"Location":"abc"}`);
  });

  it('should map independently name', () => {
    const result: number = (command as any)['mapRequestBody']({ name: 'abc' });
    assert(JSON.stringify(result) === `{"Name":"abc"}`);
  });

  it('should map independently title', () => {
    const result: number = (command as any)['mapRequestBody']({ title: 'abc' });
    assert(JSON.stringify(result) === `{"Title":"abc"}`);
  });

  it('should map independently group', () => {
    const result: number = (command as any)['mapRequestBody']({ group: 'abc' });
    assert(JSON.stringify(result) === `{"Group":"abc"}`);
  });

  it('fails validation if invalid id', () => {
    const actual = commandOptionsSchema.safeParse({ id: '1', webUrl: 'https://contoso.sharepoint.com', title: 'title' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if undefined id', () => {
    const actual = commandOptionsSchema.safeParse({ webUrl: 'https://contoso.sharepoint.com', title: 'title' });
    assert.strictEqual(actual.success, false);
  });

  it('fails if non existing PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      rights: 'abc'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `Rights option 'abc' is not recognized as valid PermissionKind choice. Please note it is case-sensitive`);
  });

  it('has correct PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      rights: 'FullMask'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails if clientSideComponentId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      clientSideComponentId: 'abc'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'ClientSideComponentId abc is not a valid GUID');
  });

  it('fails if the sequence value less than 0', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      sequence: -1
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if the id value is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      id: 'foo',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not valid. Custom action id (Guid) expected');
  });

  it('fails if the sequence value is higher than 65536', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      sequence: 65537
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if both option scriptSrc and scriptBlock specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      scriptSrc: 'abc',
      scriptBlock: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('success if location that requires group option is set, but group is also set', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions'
    });
    assert.strictEqual(actual.success, true);
  });

  it('calls the correct endpoint (url) when scope is Web', async () => {
    const postRequestSpy = initDefaultPostStubs();

    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      url: `https://contoso.sharepoint.com/_api/Web/UserCustomActions('${defaultCommandOptions.id}')`
    })));
  });

  it('calls the correct endpoint (url) when scope is Site', async () => {
    const postRequestSpy = initDefaultPostStubs();

    defaultCommandOptions.scope = "Site";
    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      url: `https://contoso.sharepoint.com/_api/Site/UserCustomActions('${defaultCommandOptions.id}')`
    })));
  });

  it('should have X-HTTP-Method MERGE header when scope is Web', async () => {
    const postRequestSpy = initDefaultPostStubs();

    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      }
    })));
  });

  it('should have X-HTTP-Method MERGE when scope is Site', async () => {
    const postRequestSpy = initDefaultPostStubs();

    defaultCommandOptions.scope = "Site";
    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      headers: {
        accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE'
      }
    })));
  });

  it('rejects invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      scope: 'All1'
    });
    assert.strictEqual(actual.success, false);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      id: '058140e3-0e37-44fc-a1d3-79c487d371a3',
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});