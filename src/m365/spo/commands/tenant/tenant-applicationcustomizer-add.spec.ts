import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './tenant-applicationcustomizer-add.js';
import { spo } from '../../../../utils/spo.js';
import { ListItemListOptions, spoListItem } from '../../../../utils/spoListItem.js';

describe(commands.TENANT_APPLICATIONCUSTOMIZER_ADD, () => {
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const customizerTitle = 'Some Customizer';
  const webTemplate = 'GROUP#0';
  const webUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `${webUrl}/sites/apps`;
  const solutionId = 'ac555cb1-e5ac-409e-86dc-61e6651b1e66';
  const clientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldApplicationCustomizer\",\"componentType\":\"Extension\",\"extensionType\":\"ApplicationCustomizer\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-application-customizer\",\"scriptResources\":{\"hello-world-application-customizer\":{\"type\":\"path\",\"path\":\"hello-world-application-customizer_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldApplicationCustomizerStrings\":{\"type\":\"path\",\"path\":\"HelloWorldApplicationCustomizerStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
  const solution = { "FileSystemObjectType": 0, "Id": 40, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "ClientComponentId": clientSideComponentId, "ClientComponentManifest": clientComponentManifest, "SolutionId": solutionId, "Created": "2022-11-03T11:25:17", "Modified": "2022-11-03T11:26:03" };
  const solutionResponse = [solution];
  const application = { "FileSystemObjectType": 0, "Id": 31, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "SkipFeatureDeployment": true, "ContainsTenantWideExtension": true, "Modified": "2022-11-03T11:26:03", "CheckoutUserId": null, "EditorId": 9 };
  const applicationResponse = [application];
  const listItemResponse = {
    Attachments: false,
    AuthorId: 3,
    ContentTypeId: '0x0100B21BD271A810EE488B570BE49963EA34',
    Created: new Date('2018-03-15T10:43:10Z'),
    EditorId: 3,
    GUID: 'ea093c7b-8ae6-4400-8b75-e2d01154dffc',
    Id: 0,
    ID: 0,
    Modified: new Date('2018-03-15T10:43:10Z'),
    Title: 'listTitle',
    RoleAssignments: []
  };

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
    auth.connection.active = true;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      spo.getTenantAppCatalogUrl,
      spoListItem.addListItem,
      spoListItem.getListItems
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_APPLICATIONCUSTOMIZER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a tenant-wide application customizer', async () => {
    let executeCommandCalled = false;

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return solutionResponse as any[];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          return applicationResponse as any[];
        }
      }

      throw 'Invalid request';
    });

    sinon.stub(spoListItem, 'addListItem').callsFake(async () => {
      executeCommandCalled = true;
      return listItemResponse;
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, title: customizerTitle, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('adds a tenant-wide application customizer to a specific webtemplate including clientSideComponentProperties', async () => {
    let executeCommandCalled = false;

    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return solutionResponse as any[];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          return applicationResponse as any[];
        }
      }

      throw 'Invalid request';
    });

    sinon.stub(spoListItem, 'addListItem').callsFake(async () => {
      executeCommandCalled = true;
      return listItemResponse;
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, title: customizerTitle, webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('throws an error when no app catalog is found', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError('Cannot add tenant-wide application customizer as app catalog cannot be found'));
  });

  it('throws an error when specific client side component is not found in manifest list', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return [];
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError('No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog'));
  });

  it('throws an error when the manifest of a specific client side component is not of type application customizer', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          const faultyClientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldApplicationCustomizer\",\"componentType\":\"Extension\",\"extensionType\":\"FormCustomizer\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-application-customizer\",\"scriptResources\":{\"hello-world-application-customizer\":{\"type\":\"path\",\"path\":\"hello-world-application-customizer_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldApplicationCustomizerStrings\":{\"type\":\"path\",\"path\":\"HelloWorldApplicationCustomizerStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
          const solutionDuplicate = { ...solution };
          solutionDuplicate.ClientComponentManifest = faultyClientComponentManifest;
          return [solutionDuplicate];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          return applicationResponse as any[];
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError(`The extension type of this component is not of type 'ApplicationCustomizer' but of type 'FormCustomizer'`));
  });

  it('throws an error when solution is not found in app catalog', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return solutionResponse as any[];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          return [];
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError(`No component found with the solution id ${solutionId}. Make sure that the solution is available in the app catalog`));
  });

  it('throws an error when solution does not contain extension that can be deployed tenant-wide', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return solutionResponse as any[];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          const faultyApplication = { ...application };
          faultyApplication.ContainsTenantWideExtension = false;
          return [faultyApplication];
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError(`The solution does not contain an extension that can be deployed to all sites. Make sure that you've entered the correct component Id.`));
  });

  it('throws an error when solution is not deployed globally', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(appCatalogUrl);
    sinon.stub(spoListItem, 'getListItems').callsFake(async (options: ListItemListOptions) => {
      if (options.webUrl === appCatalogUrl) {
        if (options.listUrl === '/sites/apps/Lists/ComponentManifests') {
          return solutionResponse as any[];
        }
        else if (options.listUrl === '/sites/apps/AppCatalog') {
          const faultyApplication = { ...application };
          faultyApplication.SkipFeatureDeployment = false;
          return [faultyApplication];
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError(`The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.`));
  });

  it('fails validation if clientSideComponentId is not a valid Guid', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all properties are specified', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});