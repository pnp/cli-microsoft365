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
const command: Command = require('./customaction-list');

describe(commands.CUSTOMACTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
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
    assert.strictEqual(command.name.startsWith(commands.CUSTOMACTION_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Name', 'Location', 'Scope', 'Id']);
  });

  it('getCustomActions called once when scope is Web', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Web'
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(getRequestSpy.calledOnce);
      assert(getCustomActionsSpy.calledWith({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Web'
      }));
      assert(getCustomActionsSpy.calledOnce);
    }
    finally {
      sinonUtil.restore((command as any)['getCustomActions']);
    }
  });

  it('getCustomActions called once when scope is Site', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site'
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(getRequestSpy.calledOnce);
      assert(getCustomActionsSpy.calledWith({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'Site'
      }));
      assert(getCustomActionsSpy.calledOnce);
    }
    finally {
      sinonUtil.restore((command as any)['getCustomActions']);
    }
  });

  it('returns all properties for output JSON', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [{ "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }, { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }] });
      }

      return Promise.reject('Invalid request');
    });

    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    };

    try {
      await command.action(logger, { options: options } as any);
      assert(loggerLogSpy.calledWith([{ "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }, { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }]));
    }
    finally {
      sinonUtil.restore((command as any)['getCustomActions']);
    }
  });

  it('getCustomActions called twice when scope is All', async () => {
    const getRequestSpy = sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    const getCustomActionsSpy = sinon.spy((command as any), 'getCustomActions');

    try {
      await command.action(logger, {
        options: {
          debug: true,
          webUrl: 'https://contoso.sharepoint.com'
        }
      });
      assert(getRequestSpy.calledTwice);
      assert(getCustomActionsSpy.calledTwice);
    }
    finally {
      sinonUtil.restore((command as any)['getCustomActions']);
    }
  });

  it('searchAllScopes called when scope is All', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve('abc');
      }

      return Promise.reject('Invalid request');
    });

    const searchAllScopesSpy = sinon.spy((command as any), 'searchAllScopes');
    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: "All"
    };

    try {
      await assert.rejects(command.action(logger, { options: options } as any));
      assert(searchAllScopesSpy.calledWith(sinon.match(
        {
          webUrl: 'https://contoso.sharepoint.com'
        })));
      assert(searchAllScopesSpy.calledOnce);
    }
    finally {
      sinonUtil.restore((command as any)['searchAllScopes']);
    }
  });

  it('searchAllScopes correctly handles no custom actions when All scope specified', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: false,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles no custom actions when All scope specified (verbose)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    });
    assert(loggerLogToStderrSpy.calledWith(`Custom actions not found`));
  });

  it('correctly handles web custom action reject request', async () => {
    const err = 'Invalid web custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      }
    }), new CommandError(err));
  });

  it('correctly handles site custom action reject request', async () => {
    const err = 'Invalid site custom action reject request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({ value: [] });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        verbose: true,
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

  it('retrieves all available user custom actions', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { webUrl: 'https://contoso.sharepoint.com/sites/abc' } });
    assert(loggerLogSpy.calledWith([
      {
        Name: 'customaction2',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Scope: 'Site',
        Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df56'
      },
      {
        Name: 'customaction1',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Scope: 'Web',
        Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df55'
      }]
    ));
  });

  it('correctly handles no scope entered (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        });
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return Promise.resolve({
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, {
      options: {
        webUrl: 'https://contoso.sharepoint.com/sites/abc',
        debug: true
      }
    });

    let correctLogStatement = false;
    log.forEach(l => {
      if (!l || typeof l !== 'string') {
        return;
      }

      if (l.indexOf('Attempt to get custom actions list with scope: All') > -1) {
        correctLogStatement = true;
      }
    });

    assert(correctLogStatement);
    assert(loggerLogSpy.calledWith([
      {
        Name: 'customaction2',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Scope: 'Site',
        Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df56'
      },
      {
        Name: 'customaction1',
        Location: 'Microsoft.SharePoint.StandardMenu',
        Scope: 'Web',
        Id: '9f7ade35-0f8d-4c8a-82e1-5008ab42df55'
      }]
    ));
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the url and scope options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com",
        scope: "Site"
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
          webUrl: "https://contoso.sharepoint.com"
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
