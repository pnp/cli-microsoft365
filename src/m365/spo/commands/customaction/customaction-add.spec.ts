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
import command, { options } from './customaction-add.js';

describe(commands.CUSTOMACTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    defaultCommandOptions = {
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions'
    };
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CUSTOMACTION_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has a permissionsKindMap', () => {
    assert.strictEqual((command as any)['permissionsKindMap'].length, 37);
  });

  it('correct https data send when custom action with location StandardMenu', async () => {
    const postRequestSpy = initDefaultPostStubs();

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        title: 'title1',
        name: 'name1',
        sequence: 102,
        location: 'Microsoft.SharePoint.StandardMenu',
        description: 'description1',
        group: 'SiteActions',
        actionUrl: '~site/Shared%20Documents/Forms/AllItems.aspx'
      })
    });
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

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        title: 'title1',
        name: 'name1',
        sequence: 101,
        location: 'ClientSideExtension.ApplicationCustomizer',
        description: 'description1',
        clientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
        clientSideComponentProperties: '{"testMessage":"Test message"}',
        debug: true
      })
    });
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

    await command.action(logger, {
      options: commandOptionsSchema.parse({
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
      })
    });
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
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 107,
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      actionUrl: '~site/SitePages/Home.aspx',
      rights: 'AddListItems,DeleteListItems,ManageLists'
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

  it('retrieves and prints the added user custom actions details when verbose specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return 'abc';
      }

      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return { FormDigestValue: 'abc' };
      }

      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return {
          "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
          "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
          "CommandUIExtension": null,
          "Description": null,
          "Group": null,
          "Id": "d26af83a-6421-4bb3-9f5c-8174ba645c80",
          "ImageUrl": null,
          "Location": "ClientSideExtension.ApplicationCustomizer",
          "Name": "{d26af83a-6421-4bb3-9f5c-8174ba645c80}",
          "RegistrationId": null,
          "RegistrationType": 0,
          "Rights": { "High": 0, "Low": 0 },
          "Scope": "1",
          "ScriptBlock": null,
          "ScriptSrc": null,
          "Sequence": 65536,
          "Title": "Places",
          "Url": null,
          "VersionOfUserCustomAction": "1.0.1.0"
        };
      }
      throw 'Invalid request';
    });

    defaultCommandOptions.verbose = true;

    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(loggerLogToStderrSpy.calledWith({
      ClientSideComponentId: '015e0fcf-fe9d-4037-95af-0a4776cdfbb4',
      ClientSideComponentProperties: '{"testMessage":"Test message"}',
      CommandUIExtension: null,
      Description: null,
      Group: null,
      Id: 'd26af83a-6421-4bb3-9f5c-8174ba645c80',
      ImageUrl: null,
      Location: 'ClientSideExtension.ApplicationCustomizer',
      Name: '{d26af83a-6421-4bb3-9f5c-8174ba645c80}',
      RegistrationId: null,
      RegistrationType: 0,
      Rights: '{"High":0,"Low":0}',
      Scope: 'Web',
      ScriptBlock: null,
      ScriptSrc: null,
      Sequence: 65536,
      Title: 'Places',
      Url: null,
      VersionOfUserCustomAction: '1.0.1.0'
    }));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) }),
      new CommandError('An error has occurred'));
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

  it('fails if non existing PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      rights: 'abc'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, `Rights option 'abc' is not recognized as valid PermissionKind choice. Please note it is case sensitive`);
  });

  it('has correct PermissionKind rights specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      rights: 'FullMask'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails if clientSideComponentId not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      clientSideComponentProperties: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if clientSideComponentId is not a valid GUID', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      clientSideComponentId: 'abc'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'ClientSideComponentId abc is not a valid GUID');
  });

  it('fails if the sequence value less than 0', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      sequence: -1
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if the sequence value is higher than 65536', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      sequence: 65537
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if both option scriptSrc and scriptBlock specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'ScriptLink',
      scriptSrc: 'abc',
      scriptBlock: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if scriptSrc or scriptBlock, but the location is not ScriptLink', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'abc',
      scriptSrc: 'abc',
      scriptBlock: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if scriptSrc and scriptBlock not specified when location ScriptLink', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'ScriptLink',
      group: 'SiteActions'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if registrationType, but not registrationId', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      registrationType: 'List'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if registrationId, but not registrationType', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      registrationId: 'abc'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if the specified URL is invalid', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('fails if location that requires group option is set, but group is not set', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails if registrationType is not a string', () => {
    const actual = commandOptionsSchema.safeParse({
      ...defaultCommandOptions,
      registrationType: 1
    });

    assert.strictEqual(actual.success, false);
  });

  it('fails if registrationType is not recognized', () => {
    const actual = commandOptionsSchema.safeParse({
      ...defaultCommandOptions,
      registrationType: 'Invalid'
    });

    assert.strictEqual(actual.success, false);
  });

  it('success if location that requires group option is set, but group is also set', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions'
    });
    assert.strictEqual(actual.success, true);
  });

  it('calls the correct endpoint (url) when scope is Web', async () => {
    const postRequestSpy = initDefaultPostStubs();

    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      url: 'https://contoso.sharepoint.com/_api/Web/UserCustomActions'
    })));
  });

  it('calls the correct endpoint (url) when scope is Site', async () => {
    const postRequestSpy = initDefaultPostStubs();

    defaultCommandOptions.scope = "Site";
    await command.action(logger, { options: commandOptionsSchema.parse(defaultCommandOptions) });
    assert(postRequestSpy.calledWith(sinon.match({
      url: 'https://contoso.sharepoint.com/_api/Site/UserCustomActions'
    })));
  });

  it('rejects invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      scope: 'All'
    });
    assert.strictEqual(actual.success, false);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('does not apply location-specific validation when location is not specified', () => {
    const schemaWithOptionalLocation = command.getRefinedSchema(
      options.partial({ location: true }) as unknown as typeof options
    );

    const actual = schemaWithOptionalLocation!.safeParse({
      webUrl: 'https://contoso.sharepoint.com',
      title: 'title',
      name: 'name',
      scriptSrc: 'https://contoso.sharepoint.com/SiteAssets/script.js'
    });

    assert.strictEqual(actual.success, true);
  });
});
