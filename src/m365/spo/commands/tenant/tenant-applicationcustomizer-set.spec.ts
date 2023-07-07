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
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as spoListItemListCommand from '../listitem/listitem-list';
import { urlUtil } from '../../../../utils/urlUtil';
import * as os from 'os';
const command: Command = require('./tenant-applicationcustomizer-set');

describe(commands.TENANT_APPLICATIONCUSTOMIZER_SET, () => {
  const title = 'Some customizer';
  const newTitle = 'New customizer';
  const newClientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed8c0bc';
  const id = 3;
  const clientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed7c0bc';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const webTemplate = "GROUP#0";
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = 'https://contoso.sharepoint.com/sites/apps';
  const solutionId = 'ac555cb1-e5ac-409e-86dc-61e6651b1e66';
  const clientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldApplicationCustomizer\",\"componentType\":\"Extension\",\"extensionType\":\"ApplicationCustomizer\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-application-customizer\",\"scriptResources\":{\"hello-world-application-customizer\":{\"type\":\"path\",\"path\":\"hello-world-application-customizer_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldApplicationCustomizerStrings\":{\"type\":\"path\",\"path\":\"HelloWorldApplicationCustomizerStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
  const solution = { "FileSystemObjectType": 0, "Id": 40, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "ClientComponentId": clientSideComponentId, "ClientComponentManifest": clientComponentManifest, "SolutionId": solutionId, "Created": "2022-11-03T11:25:17", "Modified": "2022-11-03T11:26:03" };
  const solutionResponse = [solution];
  const application = { "FileSystemObjectType": 0, "Id": 31, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "SkipFeatureDeployment": true, "ContainsTenantWideExtension": true, "Modified": "2022-11-03T11:26:03", "CheckoutUserId": null, "EditorId": 9 };
  const applicationResponse = [application];
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

  const postCallsStubClientSideComponentId = (): sinon.SinonStub => {
    return sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(3)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId
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
      cli.getSettingWithDefaultValue,
      Cli.executeCommand,
      Cli.executeCommandWithOutput
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
    const actual = await command.validate({ options: { clientSideComponentId: 'abc', newTitle: newTitle } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the newClientSideComponentId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: id, newClientSideComponentId: 'abc' } }, commandInfo);
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

  it('passes validation if newClientSideComponentId is valid', async () => {
    const actual = await command.validate({ options: { id: id, newClientSideComponentId: newClientSideComponentId } }, commandInfo);
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

  it('updates title of an application customizer by title', async () => {
    defaultGetCallStub("Title eq 'Some customizer'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    });

    assert.deepEqual(executeCallsStub.firstCall.args[0].data, { formValues: [{ FieldName: 'Title', FieldValue: 'New customizer' }] });
  });

  it('updates client side component id of an application customizer by title', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify(applicationResponse) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Title eq 'Some customizer'");
    const executeCallsStub: sinon.SinonStub = postCallsStubClientSideComponentId();
    await command.action(logger, {
      options: {
        title: title, newClientSideComponentId: newClientSideComponentId
      }
    });
    assert.deepEqual(executeCallsStub.firstCall.args[0].data, { formValues: [{ FieldName: 'TenantWideExtensionComponentId', FieldValue: '7096cded-b83d-4eab-96f0-df477ed8c0bc' }] });
  });

  it('updates properties of an application customizer by id', async () => {
    defaultGetCallStub("Id eq '3'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        id: id, clientSideComponentProperties: clientSideComponentProperties, verbose: true
      }
    });
    assert.deepEqual(executeCallsStub.firstCall.args[0].data, { formValues: [{ FieldName: 'TenantWideExtensionComponentProperties', FieldValue: '{ "someProperty": "Some value" }' }] });
  });

  it('updates an application customizer by clientSideComponentId', async () => {
    defaultGetCallStub("TenantWideExtensionComponentId eq '7096cded-b83d-4eab-96f0-df477ed7c0bc'");
    const executeCallsStub: sinon.SinonStub = defaultPostCallsStub();
    await command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId, webTemplate: webTemplate, verbose: true
      }
    });
    assert.deepEqual(executeCallsStub.firstCall.args[0].data, { formValues: [{ FieldName: 'TenantWideExtensionWebTemplate', FieldValue: 'GROUP#0' }] });
  });

  it('throws an error when specific client side component is not found in manifest list', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify([]) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Id eq '3'");
    postCallsStubClientSideComponentId();

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError('No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog'));
  });

  it('throws an error when client side component to update is not of type application customizer', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          const faultyClientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldApplicationCustomizer\",\"componentType\":\"Extension\",\"extensionType\":\"FormCustomizer\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-application-customizer\",\"scriptResources\":{\"hello-world-application-customizer\":{\"type\":\"path\",\"path\":\"hello-world-application-customizer_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldApplicationCustomizerStrings\":{\"type\":\"path\",\"path\":\"HelloWorldApplicationCustomizerStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
          const solutionDuplicate = { ...solution };
          solutionDuplicate.ClientComponentManifest = faultyClientComponentManifest;
          return { 'stdout': JSON.stringify([solutionDuplicate]) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Id eq '3'");
    postCallsStubClientSideComponentId();

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`The extension type of this component is not of type 'ApplicationCustomizer' but of type 'FormCustomizer'`));
  });

  it('throws an error when solution is not found in app catalog', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify([]) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Id eq '3'");
    postCallsStubClientSideComponentId();

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`));
  });

  it('throws an error when solution does not contain extension that can be deployed tenant-wide', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          const faultyApplication = { ...application };
          faultyApplication.ContainsTenantWideExtension = false;
          return { 'stdout': JSON.stringify([faultyApplication]) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Id eq '3'");
    postCallsStubClientSideComponentId();

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`The solution does not contain an extension that can be deployed to all sites. Make sure that you've entered the correct component Id.`));
  });

  it('throws an error when solution is not deployed globally', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          const faultyApplication = { ...application };
          faultyApplication.SkipFeatureDeployment = false;
          return { 'stdout': JSON.stringify([faultyApplication]) };
        }
      }

      throw 'Invalid request';
    });

    defaultGetCallStub("Id eq '3'");
    postCallsStubClientSideComponentId();

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.`));
  });
});