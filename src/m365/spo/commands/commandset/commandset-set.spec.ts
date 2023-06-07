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
import { formatting } from '../../../../utils/formatting';
import { session } from '../../../../utils/session';
const command: Command = require('./commandset-set');

describe(commands.COMMANDSET_SET, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  //#region Mocked Responses
  const validUrl = 'https://contoso.sharepoint.com';
  const validId = 'e7000aef-f756-4997-9420-01cc84f9ac9c';
  const validTitle = 'Commandset title';
  const validClientSideComponentId = 'b206e130-1a5b-4ae7-86a7-4f91c9924d0a';
  const validNewTitle = 'I have no inspiration whatsoever';
  const validClientSideComponentProperties = '{"testMessage":"Test message"}';
  const validListType = 'List';
  const validScope = 'Site';
  const validLocation = 'ContextMenu';
  const commandsetSingleResponse = {
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
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": "test",
        "Url": null,
        "VersionOfUserCustomAction": "16.0.1.0"
      }
    ]
  };
  const commandsetMultiResponse = {
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
        "Scope": 3,
        "ScriptBlock": null,
        "ScriptSrc": null,
        "Sequence": 0,
        "Title": "test",
        "Url": null,
        "VersionOfUserCustomAction": "16.0.1.0"
      },
      {
        "ClientSideComponentId": "b206e130-1a5b-4ae7-86a7-4f91c9924d0a",
        "ClientSideComponentProperties": "",
        "CommandUIExtension": null,
        "Description": null,
        "Group": null,
        "HostProperties": "",
        "Id": "1783725b-d5b5-4be8-973d-c6d8348e66f0",
        "ImageUrl": null,
        "Location": "ClientSideExtension.ListViewCommandSet.CommandBar",
        "Name": "{1783725b-d5b5-4be8-973d-c6d8348e66f0}",
        "RegistrationId": "100",
        "RegistrationType": 0,
        "Rights": {
          "High": 0,
          "Low": 0
        },
        "Scope": 3,
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
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.COMMANDSET_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if no other fields specified than url, id, scope', async () => {
    const actual = await command.validate({ options: { id: validId, webUrl: validUrl } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all options specified', async () => {
    const actual = await command.validate({
      options: {
        webUrl: validUrl, id: validId, newTitle: validNewTitle, listType: validListType, clientSideComponentProperties: validClientSideComponentProperties, scope: validScope, location: validLocation
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails if the specified URL is invalid', async () => {
    const actual = await command.validate({ options: { id: validId, webUrl: 'foo', newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if id, title and clientSideComponentId are provided', async () => {
    const actual = await command.validate({ options: { id: validId, title: validTitle, clientSideComponentId: validClientSideComponentProperties, webUrl: validUrl, newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid id', async () => {
    const actual = await command.validate({ options: { id: '1', webUrl: validUrl, newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid clientSideComponentId', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: '1', webUrl: validUrl, newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid listType', async () => {
    const actual = await command.validate({ options: { webUrl: validUrl, id: validId, listType: 'Invalid listType' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid scope', async () => {
    const actual = await command.validate({ options: { webUrl: validUrl, id: validId, scope: 'Invalid scope', newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if invalid location', async () => {
    const actual = await command.validate({ options: { webUrl: validUrl, id: validId, location: 'Invalid location', newTitle: validNewTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('throws error when no commandset found with option id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${validId}')`)) {
        return { "odata.null": true };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validUrl, id: validId, newTitle: validNewTitle
      }
    }), new CommandError(`No user commandsets with id '${validId}' found`));
  });

  it('throws error when no commandset found with option title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(validTitle)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`)) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validUrl, title: validTitle, newTitle: validNewTitle
      }
    }), new CommandError(`No user commandsets with title '${validTitle}' found`));
  });

  it('throws error when multiple commandsets found with option title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(validTitle)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return commandsetMultiResponse;
      }
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(validTitle)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validUrl, title: validTitle, newTitle: validNewTitle
      }
    }), new CommandError(`Multiple user commandsets with title '${validTitle}' found. Please disambiguate using IDs: ${commandsetMultiResponse.value[0].Id}, ${commandsetMultiResponse.value[1].Id}`));
  });

  it('throws error when no commandset found with option clientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(validClientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`)) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validUrl, clientSideComponentId: validClientSideComponentId, newTitle: validNewTitle
      }
    }), new CommandError(`No user commandsets with ClientSideComponentId '${validClientSideComponentId}' found`));
  });

  it('throws error when multiple commandsets found with option clientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(validClientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return commandsetMultiResponse;
      }
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(validClientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        webUrl: validUrl, clientSideComponentId: validClientSideComponentId, newTitle: validNewTitle
      }
    }), new CommandError(`Multiple user commandsets with ClientSideComponentId '${validClientSideComponentId}' found. Please disambiguate using IDs: ${commandsetMultiResponse.value[0].Id}, ${commandsetMultiResponse.value[1].Id}`));
  });

  it('updates a commandset with the id parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions(guid'${validId}')`)) {
        return commandsetSingleResponse.value[0];
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions('${validId}')`)) {
        return;
      }

      throw `Invalid request`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        verbose: true, webUrl: validUrl, id: validId, newTitle: validNewTitle, listType: validListType, clientSideComponentProperties: validClientSideComponentProperties, location: validLocation
      }
    }));
  });

  it('updates a commandset with the id parameter with scope Site', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions(guid'${validId}')`) {
        const response = commandsetSingleResponse.value[0];
        response.Scope = 2;
        return response;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if ((opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions('${validId}')`)) {
        return;
      }

      throw `Invalid request`;
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        verbose: true, webUrl: validUrl, id: validId, newTitle: validNewTitle, listType: validListType, clientSideComponentProperties: validClientSideComponentProperties, location: validLocation, scope: 'Site'
      }
    }));
  });

  it('updates a commandset with the title parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(validTitle)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return commandsetSingleResponse;
      }
      else if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(Title eq '${formatting.encodeQueryParameter(validTitle)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions('${validId}')`)) {
        return;
      }

      throw `Invalid request`;
    });

    await command.action(logger, { options: { webUrl: validUrl, title: validTitle, newTitle: validNewTitle, listType: 'Library', location: 'Both' } });

  });

  it('updates a commandset with the clientSideComponentId parameter', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/Web/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(validClientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return commandsetSingleResponse;
      }
      else if (opts.url === `https://contoso.sharepoint.com/_api/Site/UserCustomActions?$filter=(ClientSideComponentId eq guid'${formatting.encodeQueryParameter(validClientSideComponentId)}') and (startswith(Location,'ClientSideExtension.ListViewCommandSet'))`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url?.startsWith('https://contoso.sharepoint.com/_api/') && opts.url?.endsWith(`/UserCustomActions('${validId}')`)) {
        return;
      }

      throw `Invalid request`;
    });

    await command.action(logger, { options: { webUrl: validUrl, clientSideComponentId: validClientSideComponentId, newTitle: validNewTitle, listType: 'SitePages' } });
  });

  it('correctly handles API OData error', async () => {
    const error = {
      error: {
        message: `Something went wrong updating the commandset`
      }
    };

    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, { options: { webUrl: validUrl, id: validId, newTitle: validNewTitle } } as any),
      new CommandError(`Something went wrong updating the commandset`));
  });
});