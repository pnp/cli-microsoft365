import * as assert from 'assert';
import * as sinon from 'sinon';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { telemetry } from '../../../../telemetry';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { session } from '../../../../utils/session';
const command: Command = require('./tenant-commandset-get');

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
    auth.service.spoUrl = undefined;
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
    assert(loggerLogSpy.calledWith(commandSetResponse.value[0]));
  });

  it('handles error when multiple command sets with the specified title found', async () => {
    const errorMessage = `Multiple ListView Command Sets with ${title} were found. Please disambiguate (IDs): 4, 3`;
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
    }), new CommandError(errorMessage));
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
    assert(loggerLogSpy.calledWith(commandSetResponse.value[0]));
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
    assert(loggerLogSpy.calledWith(commandSetResponse.value[0]));
  });

  it('handles error when multiple command sets with the clientSideComponentId found', async () => {
    const errorMessage = `Multiple ListView Command Sets with ${clientSideComponentId} were found. Please disambiguate (IDs): 4, 3`;
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
    }), new CommandError(errorMessage));
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