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
import command from './tenant-commandset-get.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TENANT_COMMANDSET_GET, () => {
  const title = 'Some ListView Command Set';
  const id = 4;
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const commandSetResponse = {
    value:
      [{
        "FileSystemObjectType": 0,
        "ID": id,
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

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = spoUrl;
    commandInfo = cli.getCommandInfo(command);
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });
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
      request.get,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_COMMANDSET_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid number', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when all options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        id: id,
        clientSideComponentId: clientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when no options are specified', async () => {
    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and id options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        id: id
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and clientSideComponentId options are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        clientSideComponentId: clientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and clientSideComponentId options are specified', async () => {
    const actual = await command.validate({
      options: {
        id: id,
        clientSideComponentId: clientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if id is a valid number', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passed validation when title specified', async () => {
    const actual = await command.validate({ options: { title: title } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if clientSideComponentId is valid', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: clientSideComponentId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('throws error when tenant app catalog doesn\'t exist', async () => {
    const errorMessage = 'No app catalog URL found';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: null };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
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
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves a command set by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Title eq 'Some ListView Command Set'`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        title: title
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetResponse.value[0]));
  });

  it('handles error when multiple command sets with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Title eq 'Some ListView Command Set'`) {
        return {
          value:
            [
              { Title: title, Id: 4, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc' },
              { Title: title, Id: 3, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bd' }
            ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError("Multiple ListView Command Sets with Some ListView Command Set were found. Found: 3, 4."));
  });

  it('handles selecting single result when multiple command sets with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Title eq 'Some ListView Command Set'`) {
        return {
          value:
            [
              { Title: title, Id: 4, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc' },
              { Title: title, Id: 3, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bd' }
            ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(commandSetResponse.value[0]);

    await command.action(logger, {
      options: {
        title: title
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetResponse.value[0]));
  });

  it('retrieves a command set by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq 4`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: id
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetResponse.value[0]));
  });

  it('retrieves a command set by clientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(commandSetResponse.value[0]));
  });

  it('retrieves command set properties', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Id eq 4`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        id: id,
        tenantWideExtensionComponentProperties: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(JSON.parse(commandSetResponse.value[0].TenantWideExtensionComponentProperties)));
  });

  it('handles error when multiple command sets with the clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`) {
        return {
          value:
            [
              { Title: title, Id: 4, TenantWideExtensionComponentId: clientSideComponentId },
              { Title: 'Another customizer', Id: 3, TenantWideExtensionComponentId: clientSideComponentId }
            ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    }), new CommandError("Multiple ListView Command Sets with 7096cded-b83d-4eab-96f0-df477ed7c0bc were found. Found: 3, 4."));
  });

  it('handles error when specified command set not found', async () => {
    const errorMessage = 'The specified ListView Command Set was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Title eq 'Some ListView Command Set'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when listItemInstances are falsy', async () => {
    const errorMessage = 'The specified ListView Command Set was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and Title eq 'Some ListView Command Set'`) {
        return;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when retrieving command set', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=startswith(TenantWideExtensionLocation,'ClientSideExtension.ListViewCommandSet') and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    }), new CommandError(errorMessage));
  });
});