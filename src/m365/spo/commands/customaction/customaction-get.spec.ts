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
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./customaction-get');

describe(commands.CUSTOMACTION_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_GET), true);
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
        webUrl: 'https://contoso.sharepoint.com'
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
        webUrl: 'https://contoso.sharepoint.com'
      }
    }), new CommandError(`No user custom action with title 'YourAppCustomizer' found`));
  });

  it('retrieves and prints all details user custom actions by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
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

    await command.action(logger, {
      options: {
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith({
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
    }));
  });

  it('retrieves and prints all details user custom actions by title', async () => {
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
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        title: 'Places',
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert(loggerLogSpy.calledWith({
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
    }));
  });

  it('getCustomAction called once when scope is Web', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(getRequestSpy.calledOnce, 'getRequestSpy.calledOnce');
      assert(getCustomActionSpy.calledWith({
        id: 'b2307a39-e878-458b-bc90-03bc578531d6',
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }), 'getCustomActionSpy.calledWith');
      assert(getCustomActionSpy.calledOnce, 'getCustomActionSpy.calledOnce');
    }
    finally {
      sinonUtil.restore((command as any)['getCustomAction']);
    }
  });

  it('getCustomAction called once when scope is Site', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(getRequestSpy.calledOnce, 'getRequestSpy.calledOnce');
      assert(getCustomActionSpy.calledWith(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'Site'
        }), 'getCustomActionSpy.calledWith');
      assert(getCustomActionSpy.calledOnce, 'getCustomActionSpy.calledOnce');
    }
    finally {
      sinonUtil.restore((command as any)['getCustomAction']);
    }
  });

  it('getCustomAction called once when scope is All, but item found on web level', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');

    try {
      await command.action(logger, {
        options: {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com',
          scope: 'All'
        }
      });
      assert(getRequestSpy.calledOnce);
      assert(getCustomActionSpy.calledOnce);
    }
    finally {
      sinonUtil.restore((command as any)['getCustomAction']);
    }
  });

  it('getCustomAction called twice when scope is All, but item not found on web level', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionSpy = sinon.spy((command as any), 'getCustomAction');

    try {
      await command.action(logger, {
        options: {
          debug: true,
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com'
        }
      });
      assert(getRequestSpy.calledTwice);
      assert(getCustomActionSpy.calledTwice);
    }
    finally {
      sinonUtil.restore((command as any)['getCustomAction']);
    }
  });

  it('searchAllScopes called when scope is All', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      id: 'b2307a39-e878-458b-bc90-03bc578531d6',
      webUrl: 'https://contoso.sharepoint.com',
      scope: "All"
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(searchAllScopesSpy.calledWith(sinon.match(
        {
          id: 'b2307a39-e878-458b-bc90-03bc578531d6',
          webUrl: 'https://contoso.sharepoint.com'
        })), 'searchAllScopesSpy.calledWith');
      assert(searchAllScopesSpy.calledOnce, 'searchAllScopesSpy.calledOnce');
    }
    finally {
      sinonUtil.restore((command as any)['searchAllScopes']);
    }
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

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
        scope: 'All'
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('searchAllScopes correctly handles custom action odata.null when All scope specified (verbose)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions(') > -1) {
        return Promise.resolve({ "odata.null": true });
      }

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
        scope: 'All'
      }
    });
    assert(loggerLogToStderrSpy.calledWith(`Custom action with id ${actionId} not found`));
  });

  it('searchAllScopes correctly handles web custom action reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
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
        scope: 'All'
      }
    }), new CommandError(err));
  });

  it('searchAllScopes correctly handles site custom action reject request', async () => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
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
        verbose: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
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

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id option is not a valid guid', async () => {
    const actual = await command.validate({
      options:
      {
        id: "foo",
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id and url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the id, url and scope options specified', async () => {
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

  it('passes validation when the id and url option specified', async () => {
    const actual = await command.validate({
      options:
      {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('humanize scope shows correct value when scope odata is 2', () => {
    const actual = (command as any)["humanizeScope"](2);
    assert(actual === "Site");
  });

  it('humanize scope shows correct value when scope odata is 3', () => {
    const actual = (command as any)["humanizeScope"](3);
    assert(actual === "Web");
  });

  it('humanize scope shows the scope odata value when is different than 2 and 3', () => {
    const actual = (command as any)["humanizeScope"](1);
    assert(actual === "1");
  });

  it('accepts scope to be All', async () => {
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

  it('accepts scope to be Site', async () => {
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

  it('accepts scope to be Web', async () => {
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

  it('rejects invalid string scope', async () => {
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

  it('rejects invalid scope value specified as number', async () => {
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
