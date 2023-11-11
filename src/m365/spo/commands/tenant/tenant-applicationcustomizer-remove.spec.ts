import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-applicationcustomizer-remove.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TENANT_APPLICATIONCUSTOMIZER_REMOVE, () => {
  const title = 'Some customizer';
  const id = 4;
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizerResponse = {
    value:
      [{
        "FileSystemObjectType": 0,
        "ID": 4,
        "Id": 4,
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
        "GUID": '14125658-a9bc-4ddf-9c75-1b5767c9a337',
        "ComplianceAssetId": null,
        "TenantWideExtensionComponentId": clientSideComponentId,
        "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
        "TenantWideExtensionWebTemplate": null,
        "TenantWideExtensionListTemplate": 0,
        "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
        "TenantWideExtensionSequence": 0,
        "TenantWideExtensionHostProperties": null,
        "TenantWideExtensionDisabled": false
      }]
  };

  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    auth.service.spoUrl = spoUrl;
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
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
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      Cli.promptForConfirmation,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_REMOVE);
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and id options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        title: title,
        id: id
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title and clientSideComponentId options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        title: title,
        clientSideComponentId: clientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and clientSideComponentId options are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

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

  it('prompts before removing the specified tenant applicationcustomizer when force option not passed', async () => {
    await command.action(logger, {
      options: {
        id: id
      }
    });

    assert(promptIssued);
  });

  it('aborts removing the specified tenant applicationcustomizer when force option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, {
      options: {
        id: id
      }
    });
    assert(postSpy.notCalled);
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
        title: title,
        force: true
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
        title: title,
        force: true
      }
    }), new CommandError(errorMessage));
  });

  it('removes an application customizer by title (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'&$select=Id`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items(4)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        title: title
      }
    });
    assert(postSpy.called);
  });

  it('removes an application customizer by title with confirm', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'&$select=Id`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items(4)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        title: title,
        force: true
      }
    });
    assert(postSpy.called);
  });

  it('removes an application customizer by id (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Id eq '${id}'&$select=Id`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items(4)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        id: id
      }
    });
    assert(postSpy.called);
  });

  it('removes an application customizer by clientSideComponentId (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'&$select=Id`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items(4)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        verbose: true,
        clientSideComponentId: clientSideComponentId
      }
    });
    assert(postSpy.called);
  });

  it('handles error when multiple application customizers with the specified title found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'&$select=Id`) {
        return {
          value:
            [
              { Title: title, Id: id, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc' },
              { Title: title, Id: 5, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bd' }
            ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        force: true
      }
    }), new CommandError("Multiple application customizers with Some customizer were found. Found: 4, 5."));
  });

  it('handles selecting single result when multiple application customizers with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'&$select=Id`) {
        return {
          value:
            [
              { Title: title, Id: id, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc' },
              { Title: title, Id: 5, TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bd' }
            ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(applicationCustomizerResponse.value[0]);

    const postSpy = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items(4)`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: {
        title: title,
        force: true
      }
    });
    assert(postSpy.called);
  });

  it('handles error when multiple application customizers with the clientSideComponentId found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'&$select=Id`) {
        return {
          value:
            [
              { Title: title, Id: id, TenantWideExtensionComponentId: clientSideComponentId },
              { Title: 'Another customizer', Id: 5, TenantWideExtensionComponentId: clientSideComponentId }
            ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId,
        force: true
      }
    }), new CommandError("Multiple application customizers with 7096cded-b83d-4eab-96f0-df477ed7c0bc were found. Found: 4, 5."));
  });

  it('handles error when specified application customizer not found', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'&$select=Id`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        force: true
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when retrieving application customizer', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'&$select=Id`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});