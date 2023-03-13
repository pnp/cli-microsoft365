import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
const command: Command = require('./commandset-list');

describe(commands.COMMANDSET_LIST, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogSpy: sinon.SinonSpy;

  //#region Mocked Responses
  const validWebUrl = "https://contoso.sharepoint.com";
  const commandsetResponse = {
    value: [
      {
        "ClientSideComponentId": "b206e130-1a5b-4ae7-86a7-4f91c9924d0a",
        "ClientSideComponentProperties": "",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": "",
        "Id": "e7000aef-f756-4997-9420-01cc84f9ac9c",
        "ImageUrl": null,
        "Location": "ClientSideExtension.ListViewCommandSet.CommandBar",
        "Name": "{e7000aef-f756-4997-9420-01cc84f9ac9c}",
        "RegistrationId": "100",
        "RegistrationType": 0,
        "Rights": {
          "High": 0,
          "Low": 0
        },
        "Scope": 2,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": "test",
        "Url": null,
        "VersionOfUserCustomAction": "16.0.1.0"
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
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COMMANDSET_LIST);
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

  it('passes validation when the url options specified', async () => {
    const actual = await command.validate({
      options:
      {
        webUrl: "https://contoso.sharepoint.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves commandsets', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet')`)) {
        return commandsetResponse;
      }

      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet')`)) {
        return commandsetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, webUrl: validWebUrl } });
    assert(loggerLogSpy.calledWith([
      ...commandsetResponse.value,
      ...commandsetResponse.value
    ]));
  });

  it('retrieves commandsets with scope site', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet')`)) {
        return commandsetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, scope: 'Site' } });
    assert(loggerLogSpy.calledWith(commandsetResponse.value));
  });

  it('retrieves commandsets with scope web', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet')`)) {
        return commandsetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: validWebUrl, scope: 'Web' } });
    assert(loggerLogSpy.calledWith(commandsetResponse.value));
  });

  it('correctly handles API OData error', async () => {
    const error = `Something went wrong retrieving the commandset`;

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=startswith(Location,'ClientSideExtension.ListViewCommandSet')`) {
        throw error;
      }
    });

    await assert.rejects(command.action(logger, { options: { webUrl: validWebUrl, scope: 'Site' } } as any),
      new CommandError(error));
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
});
