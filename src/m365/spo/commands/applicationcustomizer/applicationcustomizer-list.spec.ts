import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './applicationcustomizer-list.js';

describe(commands.APPLICATIONCUSTOMIZER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let loggerLogSpy: sinon.SinonSpy;

  //#region Mocked Responses
  const validWebUrl = "https://contoso.sharepoint.com";
  const applicationcustomizerResponse = {
    value: [
      {
        "ClientSideComponentId": "4358e70e-ec3c-4713-beb6-39c88f7621d1",
        "ClientSideComponentProperties": "{\"listTitle\":\"News\",\"listViewTitle\":\"Published News\"}",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": "",
        "Id": "f405303c-6048-4636-9660-1b7b2cadaef9",
        "ImageUrl": null,
        "Location": "ClientSideExtension.ApplicationCustomizer",
        "Name": "{f405303c-6048-4636-9660-1b7b2cadaef9}",
        "RegistrationId": null,
        "RegistrationType": 0,
        "Rights": {
          "High": 0,
          "Low": 0
        },
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 65536,
        "Title": "NewsTicker",
        "Url": null,
        "VersionOfUserCustomAction": "1.0.1.0"
      }
    ]
  };
  //#endregion

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
    assert.strictEqual(command.name, commands.APPLICATIONCUSTOMIZER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['Name', 'Location', 'Scope', 'Id']);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: 'foo'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the scope is not a valid scope', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl, scope: 'Invalid Scope'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when a valid webUrl specified', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: "https://contoso.sharepoint.com"
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      webUrl: validWebUrl,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('retrieves applicationcustomizers', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`)) {
        return applicationcustomizerResponse;
      }

      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`)) {
        return applicationcustomizerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ verbose: true, webUrl: validWebUrl }) });
    assert(loggerLogSpy.calledWith([
      ...applicationcustomizerResponse.value,
      ...applicationcustomizerResponse.value
    ]));
  });

  it('retrieves applicationcustomizers with scope site', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`)) {
        return applicationcustomizerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: validWebUrl, scope: 'Site' }) });
    assert(loggerLogSpy.calledWith(applicationcustomizerResponse.value));
  });

  it('retrieves applicationcustomizers with scope web', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`)) {
        return applicationcustomizerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ webUrl: validWebUrl, scope: 'Web' }) });
    assert(loggerLogSpy.calledWith(applicationcustomizerResponse.value));
  });

  it('correctly handles API OData error', async () => {
    const error = `Something went wrong retrieving the applicationcustomizers`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        throw error;
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ webUrl: validWebUrl, scope: 'Site' }) }),
      new CommandError(error));
  });
});