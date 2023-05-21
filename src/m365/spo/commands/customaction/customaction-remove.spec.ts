import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./customaction-remove');

describe(commands.CUSTOMACTION_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let promptOptions: any;

  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });
  };

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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

    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });

    promptOptions = undefined;

    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => {
      if (settingName === "prompt") { return false; }
      else {
        return defaultValue;
      }
    }));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      Cli.prompt,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('handles error when multiple user custom actions with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
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
        });
      }

      return Promise.reject(`Invalid request`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'YourAppCustomizer',
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }), new CommandError(`Multiple user custom actions with title 'YourAppCustomizer' found. Please disambiguate using IDs: a70d8013-3b9f-4601-93a5-0e453ab9a1f3, 63aa745f-b4dd-4055-a4d7-d9032a0cfc59`));
  });

  it('handles error when no user custom actions with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
          value: [
          ]
        });
      }

      return Promise.reject(`Invalid request`);
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'YourAppCustomizer',
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
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
        confirm: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should prompt before removing custom action when confirmation argument not passed', async () => {
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('should abort custom action remove when prompt not confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));
    await command.action(logger, { options: { id: 'b2307a39-e878-458b-bc90-03bc578531d6', webUrl: 'https://contoso.sharepoint.com' } } as any);
    assert(postCallsSpy.notCalled);
  });

  it('should remove custom action by id when prompt confirmed', async () => {
    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/UserCustomActions?$filter=Title eq ') > -1) {
        return Promise.resolve({
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
        });
      }
      return Promise.reject('Invalid request');
    });

    const postCallsSpy: sinon.SinonStub = defaultPostCallsStub();
    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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
      confirm: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith({
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web',
        confirm: true
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
      confirm: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(postCallsSpy.calledOnce);
      assert(removeScopedCustomActionSpy.calledWith(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'Site',
          confirm: true
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
          confirm: true,
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
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });

    const removeScopedCustomActionSpy = sinon.spy((command as any), 'removeScopedCustomAction');

    try {
      await command.action(logger, {
        options: {
          debug: true,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          confirm: true
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
      confirm: true
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(searchAllScopesSpy.calledWith(sinon.match(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          confirm: true
        })), 'searchAllScopesSpy.calledWith');
      assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
    }
    finally {
      sinonUtil.restore((command as any)['searchAllScopes']);
    }
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await command.action(logger, {
      options: {
        verbose: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('should searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', async () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      // fakes remove custom action success (site collection)
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await command.action(logger, {
      options: {
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    });
    assert(loggerLogToStderrSpy.calledWith(`Custom action with id ${actionId} not found`));
  });

  it('should correctly handle custom action reject request (web)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // fakes remove custom action success (site)
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await assert.rejects(command.action(logger, {
      options: {
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
      }
    }), new CommandError(err));
  });

  it('should correctly handle custom action reject request (site)', async () => {
    const err = 'abc error';

    sinon.stub(request, 'post').callsFake((opts) => {
      // should return null to proceed with site when scope is All
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = 'b2307a39-e878-458b-bc90-03bc578531d6';

    await assert.rejects(command.action(logger, {
      options: {
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All',
        confirm: true
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
    const actual = await command.validate({ options: { webUrl: "https://contoso.sharepoint.com" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('should fail validation if the url option not specified', async () => {
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
