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
import command from './customaction-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.CUSTOMACTION_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptIssued: boolean = false;

  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return undefined;
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return undefined;
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
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

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CUSTOMACTION_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when multiple user custom actions with the specified title found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return {
          value: [
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'YourName',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'YourAppCustomizer',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'YourAppCustomizer',
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    }), new CommandError("Multiple user custom actions with title 'YourAppCustomizer' found. Found: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59."));
  });

  it('handles selecting single result when multiple custom actions sets with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === "https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=Title eq 'Places'") {
        return {
          value: [
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: 'a70d8013-3b9f-4601-93a5-0e453ab9a1f3',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'Places',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'Places',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            },
            {
              ClientSideComponentId: 'b41916e7-e69d-467f-b37f-ff8ecf8f99f2',
              ClientSideComponentProperties: "'{testMessage:Test message}'",
              CommandUIExtension: null,
              Description: null,
              Group: null,
              HostProperties: '',
              Id: '63aa745f-b4dd-4055-a4d7-d9032a0cfc59',
              ImageUrl: null,
              Location: 'ClientSideExtension.ApplicationCustomizer',
              Name: 'Places',
              RegistrationId: null,
              RegistrationType: 0,
              Rights: [Object],
              Scope: 3,
              ScriptBlock: null,
              ScriptSrc: null,
              Sequence: 0,
              Title: 'Places',
              Url: null,
              VersionOfUserCustomAction: '16.0.1.0'
            }
          ]
        };
      }

      throw `Invalid request`;
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
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
      Scope: '1',
      ScriptBlock: null,
      ScriptSrc: null,
      Sequence: 65536,
      Title: 'Places',
      Url: null,
      VersionOfUserCustomAction: '1.0.1.0'
    });

    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    try {
      await command.action(logger, { options: { title: 'Places', webUrl: 'https://contoso.sharepoint.com' } } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith(sinon.match(
        {
          title: 'Places',
          webUrl: 'https://contoso.sharepoint.com'
        })));
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('handles error when no user custom actions with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'YourAppCustomizer',
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    }), new CommandError(`No user custom action with title 'YourAppCustomizer' found`));
  });

  it('should user custom action removed successfully without prompting with confirmation argument', async () => {
    defaultPostCallsStub();

    await command.action(logger, {
      options: {
        verbose: false,
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com',
        force: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should prompt before removing custom action when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } });

    assert(promptIssued);
  });

  it('should abort custom action remove when prompt not confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } } as any);
    assert(postCallsSpy.notCalled);
  });

  it('should remove custom action by id when prompt confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    try {
      await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith(sinon.match(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com'
        })));
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should remove custom action by title when prompt confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return {
          value: [
            {
              "ClientSideComponentId": "015e0fcf-fe9d-4037-95af-0a4776cdfbb4",
              "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}",
              "CommandUIExtension": null,
              "Description": null,
              "Group": null,
              "Id": "b2307a39-e878-458b-bc90-03bc578531d6",
              "ImageUrl": null,
              "Location": "ClientSideExtension.ApplicationCustomizer",
              "Name": "{b2307a39-e878-458b-bc90-03bc578531d6}",
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
          ]
        };
      }
      throw 'Invalid request';
    });

    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    try {
      await command.action(logger, { options: { title: 'Places', webUrl: 'https://contoso.sharepoint.com' } } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith(sinon.match(
        {
          title: 'Places',
          webUrl: 'https://contoso.sharepoint.com'
        })));
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should removeScopedCustomAction be called once when scope is Web', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web',
      force: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith({
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web',
        force: true
      }), 'removeScopedCustomActionSpy data error');
      assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should removeScopedCustomAction be called once when scope is Site', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site',
      force: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'Site',
          force: true
        }), 'removeScopedCustomActionSpy data error');
      assert(removeScopedCustomActionSpy.calledOnce, 'removeScopedCustomActionSpy calledOnce error');
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should removeScopedCustomAction be called once when scope is All, but item found on web level', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    try {
      await command.action(logger, {
        options: {
          force: true,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'All'
        }
      });
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledOnce);
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should removeScopedCustomAction be called twice when scope is All, but item not found on web level', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return undefined;
      }

      throw 'Invalid request';
    });

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    try {
      await command.action(logger, {
        options: {
          debug: true,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          force: true
        }
      });
      assert(removeScopedCustomActionSpy.calledTwice);
    }
    finally {
      sinonUtil.restore((command as any)['removeScopedCustomAction']);
    }
  });

  it('should searchAllScopes be called when scope is All', async () => {
    defaultPostCallsStub();

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      force: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(searchAllScopesSpy.calledWith(sinon.match(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          force: true
        })), 'searchAllScopesSpy.calledWith');
      assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
    }
    finally {
      sinonUtil.restore((command as any)['searchAllScopes']);
    }
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await command.action(logger, {
      options: {
        verbose: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      }
    });
    assert(loggerLogToStderrSpy.calledWith(`Custom action with id ${actionId} not found`));
  });

  it('should correctly handle custom action reject request (web)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await assert.rejects(command.action(logger, {
      options: {
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      }
    }), new CommandError(err));
  });

  it('should correctly handle custom action reject request (site)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake(async (opts) => {
      // should return null to proceed with site when scope is All
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
      options: {
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        force: true
      }
    }), new CommandError(err));
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

  it('should fail validation if the id option not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option not specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({ options: { id: "BC448D63-484F-49C5-AB8C-96B14AA68D50" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the id option is not a valid guid', async () => {
    const actual = await command.validate({
      options:
      {
        id: "foo",
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should pass validation when the id and url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the id, url and scope options specified', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: "Site"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should pass validation when the id and url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be All', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: 'All'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Site', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should accept scope to be Web', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: 'Web'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('should reject invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('should reject invalid scope value specified as number', async () => {
    const scope = 123;
    const actual = await command.validate({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid custom action scope. Allowed values are Site|Web|All`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', async () => {
    const actual = await command.validate(
      {
        options:
        {
          id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
          webUrl: "https://contoso.sharepoint.com"
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
