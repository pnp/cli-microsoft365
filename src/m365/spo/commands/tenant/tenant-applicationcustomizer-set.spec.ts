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
import * as os from 'os';
import { session } from '../../../../utils/session';
const command: Command = require('./tenant-applicationcustomizer-set');

describe(commands.TENANT_APPLICATIONCUSTOMIZER_SET, () => {
  const title = 'Some customizer';
  const newTitle = 'New customizer';
  const id = 3;
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const webTemplate = "GROUP#0";
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const applicationCustomizerResponse = {
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
  const multipleResponses = {
    value:
      [
        { Title: title, Id: 3, TenantWideExtensionComponentId: clientSideComponentId },
        { Title: title, Id: 4, TenantWideExtensionComponentId: clientSideComponentId }
      ]
  };
  let cli: Cli;
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  const defaultGetCallStub = (filter: string): sinon.SinonStub => {
    return sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and ${filter}`) {
        return applicationCustomizerResponse;
      }

      throw 'Invalid request';
    });
  };

  const defaultPostCallsStub = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(3)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              FieldName: "Title",
              FieldValue: title
            }
          ]
        };
      }

      throw 'Invalid request';
    });
  };

  before(() => {
    cli = Cli.getInstance();
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_SET);
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
    const actual = await command.validate({
      options: {
        title: title,
        id: id,
        clientSideComponentId: clientSideComponentId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when all options are empty', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if clientSideComponentId is valid', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: clientSideComponentId, newTitle: newTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('handles error when tenant app catalog doesn\'t exist', async () => {
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
        newTitle: newTitle
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when no application customizer with the specified title found', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when multiple application customizers with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`) {
        return multipleResponses;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    }), new CommandError(`Multiple application customizers with title '${title}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponses.value.map(item => `- ${(item as any).Id}`).join(os.EOL)}`));
  });

  it('handles error when multiple application customizers with the clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`) {
        return multipleResponses;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId, newTitle: newTitle
      }
    }), new CommandError(`Multiple application customizers with ClientSideComponentId '${clientSideComponentId}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponses.value.map(item => `- ${(item as any).Id}`).join(os.EOL)}`));
  });

  it('handles error when listItemInstances are falsy', async () => {
    const errorMessage = 'The specified application customizer was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and Title eq 'Some customizer'`) {
        return { value: undefined };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when executing command', async () => {
    const errorMessage = 'An error has occurred';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/items?$filter=TenantWideExtensionLocation eq 'ClientSideExtension.ApplicationCustomizer' and TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'`) {
        throw errorMessage;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId, newTitle: newTitle
      }
    }), new CommandError(errorMessage));
  });

  it('Updates an application customizer by title', async () => {
    defaultGetCallStub("Title eq 'Some customizer'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    });
    assert(executeCallsStub.calledOnce);
  });

  it('Updates an application customizer by id', async () => {
    defaultGetCallStub("Id eq '3'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        id: id, clientSideComponentProperties: clientSideComponentProperties, verbose: true
      }
    });
    assert(executeCallsStub.calledOnce);
  });

  it('Updates an application customizer by clientSideComponentId', async () => {
    defaultGetCallStub("TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId, webTemplate: webTemplate, verbose: true
      }
    });
    assert(executeCallsStub.calledOnce);
  });
});