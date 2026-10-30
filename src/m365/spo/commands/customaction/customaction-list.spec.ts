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
import command, { options } from './customaction-list.js';

describe(commands.CUSTOMACTION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CUSTOMACTION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Name', 'Location', 'Scope', 'Id']);
  });

  it('returns all properties for output JSON', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return { value: [{ "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }, { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }] };
      }

      throw 'Invalid request';
    });

    const options = {
      webUrl: 'https://contoso.sharepoint.com',
      scope: 'Site',
      output: 'json'
    };

    await command.action(logger, { options: commandOptionsSchema.parse(options) });
    assert(loggerLogSpy.calledWith([{ "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "8b86123a-3194-49cf-b167-c044b613a48a", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }, { "ClientSideComponentId": "b41916e7-e69d-467f-b37f-ff8ecf8f99f2", "ClientSideComponentProperties": "{\"testMessage\":\"Test message\"}", "CommandUIExtension": null, "Description": null, "Group": null, "Id": "9115bb61-d9f1-4ed4-b7b7-e5d1834e60f5", "ImageUrl": null, "Location": "ClientSideExtension.ApplicationCustomizer", "Name": "YourName", "RegistrationId": null, "RegistrationType": 0, "Rights": { "High": "0", "Low": "0" }, "Scope": 3, "ScriptBlock": null, "ScriptSrc": null, "Sequence": 0, "Title": "YourAppCustomizer", "Url": null, "VersionOfUserCustomAction": "16.0.1.0" }]));
  });


  it('return an empty array when All scope specified (verbose)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return { value: [] };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      })
    });
    assert(loggerLogSpy.calledWith([]));
  });

  it('correctly handles web custom action reject request', async () => {
    const err = 'Invalid web custom action reject request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        throw err;
      }
      else if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      })
    }), new CommandError(err));
  });

  it('correctly handles site custom action reject request', async () => {
    const err = 'Invalid site custom action reject request';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return { value: [] };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        throw err;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com',
        scope: 'All'
      })
    }), new CommandError(err));
  });

  it('retrieves all available user custom actions', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return {
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return {
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: 'https://contoso.sharepoint.com/sites/abc' }) });
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_api/Web/UserCustomActions') > -1) {
        return {
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df55",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction1",
              Scope: 3
            }]
        };
      }

      if ((opts.url as string).indexOf('/_api/Site/UserCustomActions') > -1) {
        return {
          value: [
            {
              Id: "9f7ade35-0f8d-4c8a-82e1-5008ab42df56",
              Location: "Microsoft.SharePoint.StandardMenu",
              Name: "customaction2",
              Scope: 2
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        webUrl: 'https://contoso.sharepoint.com/sites/abc',
        debug: true
      })
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

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error.issues[0].message, 'foo is not a valid SharePoint Online site URL.');
  });

  it('passes validation when the url options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when the url and scope options specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: "Site"
    });
    assert.strictEqual(actual.success, true);
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

  it('accepts scope to be All', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'All'
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Site', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Site'
    });
    assert.strictEqual(actual.success, true);
  });

  it('accepts scope to be Web', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'Web'
    });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid string scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 'foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('rejects invalid scope value specified as number', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      scope: 123
    });
    assert.strictEqual(actual.success, false);
  });

  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com",
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
