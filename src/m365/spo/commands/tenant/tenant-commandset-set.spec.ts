import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as spoListItemSetCommand from '../listitem/listitem-set';
import request from '../../../../request';
const command: Command = require('./tenant-commandset-set');

describe(commands.TENANT_COMMANDSET_SET, () => {
  const id = 1;
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const title = 'Some Command Set';
  const webTemplate = 'GROUP#0';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `https://contoso.sharepoint.com/sites/apps`;
  const commandSetResponse = {
    value:
      [{
        "FileSystemObjectType": 0,
        "Id": id,
        "ServerRedirectedEmbedUri": null,
        "ServerRedirectedEmbedUrl": "",
        "ContentTypeId": "0x00693E2C487575B448BD420C12CEAE7EFE",
        "Title": title,
        "Modified": "2023-01-11T15:47:38Z",
        "Created": "2023-01-11T15:47:38Z",
        "AuthorId": 9,
        "EditorId": 9,
        "OData__UIVersionString": "1.0",
        "Attachments": false,
        "GUID": "6e6f2429-cdec-4b90-89da-139d2665919e",
        "ComplianceAssetId": null,
        "TenantWideExtensionComponentId": clientSideComponentId,
        "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
        "TenantWideExtensionWebTemplate": null,
        "TenantWideExtensionListTemplate": 101,
        "TenantWideExtensionLocation": "ClientSideExtension.ListViewCommandSet.ContextMenu",
        "TenantWideExtensionSequence": 0,
        "TenantWideExtensionHostProperties": null,
        "TenantWideExtensionDisabled": false
      }]
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    auth.service.connected = true;
    auth.service.spoUrl = spoUrl;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommand,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_COMMANDSET_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates a tenant-wide ListView Command Set for lists', async () => {
    let executeCommandCalled = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemSetCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, newTitle: title, clientSideComponentId: clientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('updates a tenant-wide ListView Command Set for lists with location ContextMenu and listType SitePages', async () => {
    let executeCommandCalled = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemSetCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, location: 'ContextMenu', listType: 'SitePages' } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('updates a tenant-wide ListView Command Set for lists with location CommandBar and listType Library ', async () => {
    let executeCommandCalled = false;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemSetCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: id, location: 'CommandBar', listType: 'Library' } });
    assert.strictEqual(executeCommandCalled, true);
  });

  const errorMessage = 'No app catalog URL found';

  it('throws error when tenant app catalog doesn\'t exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: null };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('throws error when retrieving a tenant app catalog fails with an exception', async () => {
    const errorMessage = 'Couldn\'t retrieve tenant app catalog URL';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when specified command set not found', async () => {
    const errorMessage = 'The specified command set was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('fails validation if no option to update is specified is not a valid Guid', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if clientSideComponentId is not a valid Guid', async () => {
    const actual = await command.validate({ options: { id: id, clientSideComponentId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if location value is not valid', async () => {
    const actual = await command.validate({ options: { id: id, location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listType value is not valid', async () => {
    const actual = await command.validate({ options: { id: id, listType: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all properties are specified', async () => {
    const actual = await command.validate({ options: { id: id, newTitle: title, clientSideComponentId: clientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});