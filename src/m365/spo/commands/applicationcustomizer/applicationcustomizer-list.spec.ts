import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./applicationcustomizer-list');

describe(commands.APPLICATIONCUSTOMIZER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
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

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the scope is not a valid scope', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: validWebUrl, scope: 'Invalid Scope'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when a valid webUrl specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
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

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl } });
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

    await command.action(logger, { options: { webUrl: validWebUrl, scope: 'Site' } });
    assert(loggerLogSpy.calledWith(applicationcustomizerResponse.value));
  });

  it('retrieves applicationcustomizers with scope web', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`)) {
        return applicationcustomizerResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, scope: 'Web' } });
    assert(loggerLogSpy.calledWith(applicationcustomizerResponse.value));
  });

  it('correctly handles API OData error', async () => {
    const error = `Something went wrong retrieving the applicationcustomizers`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=Location eq 'ClientSideExtension.ApplicationCustomizer'`) {
        throw error;
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: validWebUrl, scope: 'Site' } } as any),
      new CommandError(error));
  });
});