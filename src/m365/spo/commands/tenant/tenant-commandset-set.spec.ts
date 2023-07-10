import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
import * as spoListItemListCommand from '../listitem/listitem-list';
import { urlUtil } from '../../../../utils/urlUtil';
import request from '../../../../request';
import * as os from 'os';
const command: Command = require('./tenant-commandset-set');

describe(commands.TENANT_COMMANDSET_SET, () => {
  const title = 'Some Command Set';
  const id = 1;
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const newTitle = 'New Command Set';
  const newClientSideComponentId = '7096cded-b83d-4eab-96f0-df477ed8c0bc';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const webTemplate = 'GROUP#0';
  const spoUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `https://contoso.sharepoint.com/sites/apps`;
  const solutionId = 'ac555cb1-e5ac-409e-86dc-61e6651b1e66';
  const clientComponentManifest = "{\"id\":\"8645b5d7-51e1-40af-8a62-d64af8a8487e\",\"alias\":\"HelloWorldCommandSet\",\"componentType\":\"Extension\",\"extensionType\":\"ListViewCommandSet\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"items\":{\"COMMAND_1\":{\"title\":{\"default\":\"Command One\"},\"iconImageUrl\":\"icons/request.png\",\"type\":\"command\"},\"COMMAND_2\":{\"title\":{\"default\":\"Command Two\"},\"iconImageUrl\":\"icons/cancel.png\",\"type\":\"command\"}},\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-command-set\",\"scriptResources\":{\"hello-world-command-set\":{\"type\":\"path\",\"path\":\"hello-world-command-set_087bd3f44cb1a4a2316f.js\"},\"@microsoft/sp-dialog\":{\"type\":\"component\",\"id\":\"c0c518b8-701b-4f6f-956d-5782772bb731\",\"version\":\"1.17.3\"},\"@microsoft/sp-listview-extensibility\":{\"type\":\"component\",\"id\":\"d37b65ee-c7d8-4570-bc74-2b294ff3b380\",\"version\":\"1.17.3\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.17.3\"}}},\"mpnId\":\"Undefined-1.17.3\",\"clientComponentDeveloper\":\"\"}";
  const solution = { "FileSystemObjectType": 0, "Id": 67, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "ClientComponentId": clientSideComponentId, "SolutionId": solutionId, "ClientComponentManifest": clientComponentManifest, "Created": "2023-07-09T20:23:37", "Modified": "2023-07-09T20:23:37" };
  const solutionResponse = [solution];
  const application = { "FileSystemObjectType": 0, "Id": 61, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "SkipFeatureDeployment": true, "ContainsTenantWideExtension": true, "Modified": '2023-07-09T20:23:36', "CheckoutUserId": null, "EditorId": 9 };
  const applicationResponse = [application];
  const commandSetResponse = {
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
    "TenantWideExtensionComponentProperties": "{\"sampleTextOne\":\"One item is selected in the list.\", \"sampleTextTwo\":\"This command is always visible.\"}",
    "TenantWideExtensionWebTemplate": null,
    "TenantWideExtensionListTemplate": 101,
    "TenantWideExtensionLocation": "ClientSideExtension.ListViewCommandSet.CommandBar",
    "TenantWideExtensionSequence": 0,
    "TenantWideExtensionHostProperties": null,
    "TenantWideExtensionDisabled": false
  };
  const multipleResponses = {
    value:
      [
        { Title: title, Id: 3, TenantWideExtensionComponentId: clientSideComponentId },
        { Title: title, Id: 4, TenantWideExtensionComponentId: clientSideComponentId }
      ]
  };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
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
    assert.strictEqual(command.name, commands.TENANT_COMMANDSET_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates a tenant-wide listview command set for lists by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "Title",
              FieldValue: title,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 100,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionWebTemplate",
              FieldValue: webTemplate,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentProperties",
              FieldValue: clientSideComponentProperties,
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, newTitle: title, newClientSideComponentId: newClientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties, verbose: true } }));
  });

  it('updates a tenant-wide listview command set for lists with location ContextMenu and listType SitePages by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 119,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.ContextMenu",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, location: 'ContextMenu', listType: 'SitePages' } }));
  });

  it('updates a tenant-wide listview command set for lists with location CommandBar and listType Library by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionListTemplate",
              FieldValue: 101,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.CommandBar",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { id: id, location: 'CommandBar', listType: 'Library' } }));
  });

  it('updates title of tenant-wide listview command set for lists by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Title eq 'Some Command Set'`) {
        return {
          value: [
            commandSetResponse
          ]
        };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.CommandBar",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { title: title, newClientSideComponentId: newClientSideComponentId, verbose: true } }));
  });

  it('updates title of tenant-wide listview command set for lists by ClientSideComponentId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and TenantWideExtensionComponentId eq '9748c81b-d72e-4048-886a-e98649543743'`) {
        return {
          value: [
            commandSetResponse
          ]
        };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

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

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet.CommandBar",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.doesNotReject(command.action(logger, { options: { clientSideComponentId: clientSideComponentId, newClientSideComponentId: newClientSideComponentId, verbose: true } }));
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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError('No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog'));
  });

  it('throws an error when client side component to update is not of type listview commandset', async () => {
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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`The extension type of this component is not of type 'ListViewCommandSet' but of type 'FormCustomizer'`));
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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

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

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Id eq '1'`) {
        return { value: [commandSetResponse] };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)`) {
        return commandSetResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items(1)/ValidateUpdateListItem()`) {
        return {
          value: [
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionComponentId",
              FieldValue: newClientSideComponentId,
              HasException: false,
              ItemId: 4
            },
            {
              ErrorCode: 0,
              ErrorMessage: null,
              FieldName: "TenantWideExtensionLocation",
              FieldValue: "ClientSideExtension.ListViewCommandSet",
              HasException: false,
              ItemId: 4
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: id, newClientSideComponentId: newClientSideComponentId, verbose: true } }),
      new CommandError(`The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.`));
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
        id: id,
        newTitle: title
      }
    }), new CommandError(errorMessage));
  });

  it('handles error when retrieving a tenant app catalog fails with an exception', async () => {
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

  it('handles error when no listview commandset with the specified title found', async () => {
    const errorMessage = 'The specified listview commandset was not found';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Title eq 'Some Command Set'`) {
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

  it('handles error when multiple listview commandsets with the specified title found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and Title eq 'Some Command Set'`) {
        return multipleResponses;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title, newTitle: newTitle
      }
    }), new CommandError(`Multiple listview commandsets with title '${title}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponses.value.map(item => `- ${(item as any).Id}`).join(os.EOL)}`));
  });

  it('handles error when multiple listview commandsets with the clientSideComponentId found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `${spoUrl}/_api/SP_TenantSettings_Current`) {
        return { CorporateCatalogUrl: appCatalogUrl };
      }

      if (opts.url === `${spoUrl}/sites/apps/_api/web/GetList('%2Fsites%2Fapps%2Flists%2FTenantWideExtensions')/Items?$filter=startswith(TenantWideExtensionLocation, 'ClientSideExtension.ListViewCommandSet') and TenantWideExtensionComponentId eq '9748c81b-d72e-4048-886a-e98649543743'`) {
        return multipleResponses;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        clientSideComponentId: clientSideComponentId, newTitle: newTitle
      }
    }), new CommandError(`Multiple listview commandsets with ClientSideComponentId '${clientSideComponentId}' found. Please disambiguate using IDs: ${os.EOL}${multipleResponses.value.map(item => `- ${(item as any).Id}`).join(os.EOL)}`));
  });

  it('fails validation if the id is not a number', async () => {
    const actual = await command.validate({ options: { id: 'abc', newTitle: newTitle } }, commandInfo);
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

  it('fails validation if no option to update is specified', async () => {
    const actual = await command.validate({ options: { id: id } }, commandInfo);
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

  it('passes validation if clientSideComponentId is valid', async () => {
    const actual = await command.validate({ options: { clientSideComponentId: clientSideComponentId, newTitle: newTitle } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if newClientSideComponentId is valid', async () => {
    const actual = await command.validate({ options: { id: id, newClientSideComponentId: newClientSideComponentId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all properties are specified', async () => {
    const actual = await command.validate({ options: { id: id, newTitle: title, newClientSideComponentId: newClientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});