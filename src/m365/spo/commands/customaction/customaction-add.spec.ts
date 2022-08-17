import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./customaction-add');

describe(commands.CUSTOMACTION_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let defaultCommandOptions: any;
  const initDefaultPostStubs = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });
  };

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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    defaultCommandOptions = {
      url: 'https://contoso.sharepoint.com',
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
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('has a permissionsKindMap', () => {
    assert.strictEqual((command as any)['permissionsKindMap'].length, 37);
  });

  it('correct https data send when custom action with location StandardMenu', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with location ClientSideExtension.ApplicationCustomizer', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with location ClientSideExtension.ListViewCommandSet', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with location EditControlBlock', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with location ScriptLink', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with location ScriptLink and ScriptBlock', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 108,
      location: 'ScriptLink',
      scriptBlock: '(function(){ return console.log("Hello CLI for Microsoft 365!"); })();'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          data: {
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

  it('correct https data send when custom action with location CommandUI.Ribbon', (done) => {
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
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correct https data send when custom action with delegated rights', (done) => {
    const postRequestSpy = initDefaultPostStubs();
    const options: any = {
      url: 'https://contoso.sharepoint.com',
      title: 'title1',
      name: 'name1',
      sequence: 107,
      location: 'Microsoft.SharePoint.StandardMenu',
      group: 'SiteActions',
      actionUrl: '~site/SitePages/Home.aspx',
      rights: 'AddListItems,DeleteListItems,ManageLists'
    };

    command.action(logger, { options: options } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves and prints the added user custom actions details when verbose specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if ((opts.url as string).indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve(
          {
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
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    defaultCommandOptions.verbose = true;

    command.action(logger, { options: defaultCommandOptions } as any, () => {
      try {
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
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, { options: defaultCommandOptions } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('offers autocomplete for the registrationType option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--registrationType') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the rights option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--rights') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
  });

  it('offers autocomplete for the scope option', () => {
    const options = command.options;
    for (let i = 0; i < options.length; i++) {
      if (options[i].option.indexOf('--scope') > -1) {
        assert(options[i].autocomplete);
        return;
      }
    }
    assert(false);
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

  it('fails if non existing PermissionKind rights specified', async () => {
    defaultCommandOptions.rights = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Rights option '${defaultCommandOptions.rights}' is not recognized as valid PermissionKind choice. Please note it is case sensitive`);
  });

  it('has correct PermissionKind rights specified', async () => {
    defaultCommandOptions.rights = 'FullMask';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert(actual === true);
  });

  it('fails if clientSideComponentId not specified', async () => {
    defaultCommandOptions.clientSideComponentProperties = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Option clientSideComponentProperties is specified, but the clientSideComponentId option is missing`);
  });

  it('fails if clientSideComponentId is not a valid GUID', async () => {
    defaultCommandOptions.clientSideComponentId = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `ClientSideComponentId ${defaultCommandOptions.clientSideComponentId} is not a valid GUID`);
  });

  it('fails if the sequence value less than 0', async () => {
    defaultCommandOptions.sequence = -1;
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Invalid option sequence. Expected value in range from 0 to 65536`);
  });

  it('fails if the sequence value is higher than 65536', async () => {
    defaultCommandOptions.sequence = 65537;
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Invalid option sequence. Expected value in range from 0 to 65536`);
  });

  it('fails if both option scriptSrc and scriptBlock specified', async () => {
    defaultCommandOptions.location = 'ScriptLink';
    defaultCommandOptions.scriptSrc = 'abc';
    defaultCommandOptions.scriptBlock = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Either option scriptSrc or scriptBlock can be specified, but not both`);
  });

  it('fails if scriptSrc or scriptBlock, but the location is not ScriptLink', async () => {
    defaultCommandOptions.location = 'abc';
    defaultCommandOptions.scriptSrc = 'abc';
    defaultCommandOptions.scriptBlock = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Option scriptSrc or scriptBlock is specified, but the location option is different than ScriptLink. Please use --actionUrl, if the location should be different than ScriptLink`);
  });

  it('fails if scriptSrc and scriptBlock not specified when location ScriptLink', async () => {
    defaultCommandOptions.location = 'ScriptLink';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Option scriptSrc or scriptBlock is required when the location is set to ScriptLink`);
  });

  it('fails if registrationType, but not registrationId', async () => {
    defaultCommandOptions.registrationType = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Option registrationType is specified, but registrationId is missing`);
  });

  it('fails if registrationId, but not registrationType', async () => {
    defaultCommandOptions.registrationId = 'abc';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `Option registrationId is specified, but registrationType is missing`);
  });

  it('fails if the specified URL is invalid', async () => {
    defaultCommandOptions.location = 'Microsoft.SharePoint.StandardMenu';
    defaultCommandOptions.group = 'SiteActions';
    defaultCommandOptions.url = 'foo';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails if location that requires group option is set, but group is not set', async () => {
    defaultCommandOptions.location = 'Microsoft.SharePoint.StandardMenu';
    defaultCommandOptions.group = undefined;
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert.strictEqual(actual, `The location specified requires the group option to be specified as well`);
  });

  it('success if location that requires group option is set, but group is also set', async () => {
    defaultCommandOptions.location = 'Microsoft.SharePoint.StandardMenu';
    defaultCommandOptions.group = 'SiteActions';
    const actual = await command.validate({ options: defaultCommandOptions }, commandInfo);
    assert(actual === true);
  });

  it('calls the correct endpoint (url) when scope is Web', (done) => {
    const postRequestSpy = initDefaultPostStubs();

    command.action(logger, { options: defaultCommandOptions } as any, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          url: 'https://contoso.sharepoint.com/_api/Web/UserCustomActions'
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
    command.action(logger, { options: defaultCommandOptions } as any, () => {
      try {
        assert(postRequestSpy.calledWith(sinon.match({
          url: 'https://contoso.sharepoint.com/_api/Site/UserCustomActions'
        })));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('rejects invalid string scope', async () => {
    defaultCommandOptions.scope = 'All';
    const actual = await command.validate({
      options: defaultCommandOptions
    }, commandInfo);
    assert.strictEqual(actual, `${defaultCommandOptions.scope} is not a valid custom action scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', async () => {
    const actual = await command.validate(
      {
        options: defaultCommandOptions
      }, commandInfo);
    assert(actual === true);
  });

  it('supports specifying scope', () => {
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
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
});