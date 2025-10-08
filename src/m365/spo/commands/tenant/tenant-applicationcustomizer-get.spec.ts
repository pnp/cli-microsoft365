import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-applicationcustomizer-get.js';
import { settingsNames } from '../../../../settingsNames.js';
import { spo } from '../../../../utils/spo.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';

describe(commands.TENANT_APPLICATIONCUSTOMIZER_GET, () => {
  const title = 'Some customizer';
  const id = 4;
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizerResponse = [{
    "FileSystemObjectType": 0,
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
    "GUID": id,
    "ComplianceAssetId": null,
    "TenantWideExtensionComponentId": clientSideComponentId,
    "TenantWideExtensionComponentProperties": "{\"testMessage\":\"Test message\"}",
    "TenantWideExtensionWebTemplate": null,
    "TenantWideExtensionListTemplate": 0,
    "TenantWideExtensionLocation": "ClientSideExtension.ApplicationCustomizer",
    "TenantWideExtensionSequence": 0,
    "TenantWideExtensionHostProperties": null,
    "TenantWideExtensionDisabled": false
  }];

  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = spoUrl;
    commandInfo = cli.getCommandInfo(command);
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
      spo.getTenantAppCatalogUrl,
      spoListItem.getListItems,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a number', async () => {
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

  it('passes validation if id is a valid GUID', async () => {
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
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('throws error when retrieving a tenant app catalog fails with an exception', async () => {
    const errorMessage = 'Couldn\'t retrieve tenant app catalog URL';

    sinon.stub(spo, 'getTenantAppCatalogUrl').callsFake(() => {
      throw errorMessage;
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('retrieves an application customizer by title', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`
        ) {
          return applicationCustomizerResponse as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, {
      options: {
        title: title
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(applicationCustomizerResponse[0]));
  });

  it('handles error when multiple application customizers with the specified title found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`
        ) {
          return [
            { Title: title, GUID: '14125658-a9bc-4ddf-9c75-1b5767c9a337', TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bc' },
            { Title: title, GUID: '14125658-a9bc-4ddf-9c75-1b5767c9a338', TenantWideExtensionComponentId: '7096cded-b83d-4eab-96f0-df477ed7c0bd' }
          ] as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError("Multiple application customizers with Some customizer were found. Found: undefined."));
  });

  it('handles selecting single result when multiple application customizers with the specified name found and cli is set to prompt', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`
        ) {
          return [
            applicationCustomizerResponse[0],
            applicationCustomizerResponse[0]
          ] as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(applicationCustomizerResponse[0]);

    await command.action(logger, {
      options: {
        title: title
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(applicationCustomizerResponse[0]));
  });

  it('retrieves an application customizer by id', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Id eq 4`
        ) {
          return applicationCustomizerResponse as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, {
      options: {
        id: id
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(applicationCustomizerResponse[0]));
  });

  it('retrieves an application customizer by clientSideComponentId', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`
        ) {
          return applicationCustomizerResponse as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(applicationCustomizerResponse[0]));
  });

  it('retrieves an application customizer component properties', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Id eq 4`
        ) {
          return applicationCustomizerResponse as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await command.action(logger, {
      options: {
        id: id,
        tenantWideExtensionComponentProperties: true
      }
    });
    assert(loggerLogSpy.calledOnceWithExactly(JSON.parse(applicationCustomizerResponse[0].TenantWideExtensionComponentProperties)));
  });

  it('handles error when multiple application customizers with the clientSideComponentId found', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`
        ) {
          return [
            { Title: title, GUID: '14125658-a9bc-4ddf-9c75-1b5767c9a337', TenantWideExtensionComponentId: clientSideComponentId },
            { Title: 'Another customizer', GUID: '14125658-a9bc-4ddf-9c75-1b5767c9a338', TenantWideExtensionComponentId: clientSideComponentId }
          ] as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    }), new CommandError("Multiple application customizers with 7096cded-b83d-4eab-96f0-df477ed7c0bc were found. Found: undefined."));
  });

  it('handles error when specified application customizer not found', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`
        ) {
          return [];
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when listItemInstances are falsy', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`
        ) {
          return undefined as any;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when retrieving application customizer', async () => {
    const errorMessage = 'An error has occurred';

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === `/Lists/TenantWideExtensions` &&
          options.filter === `TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`
        ) {
          throw errorMessage;
        }
      }

      throw 'Invalid request: ' + JSON.stringify(options);
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId
      }
    }), new CommandError(errorMessage));
  });
});