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
import * as spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
import * as spoListItemListCommand from '../listitem/listitem-list';
import * as spoListItemAddCommand from '../listitem/listitem-add';
import { urlUtil } from '../../../../utils/urlUtil';
const command: Command = require('./tenant-commandset-add');

describe(commands.TENANT_COMMANDSET_ADD, () => {
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const clientSideComponentProperties = '{ "someProperty": "Some value" }';
  const customizerTitle = 'Some Command Set';
  const webTemplate = 'GROUP#0';
  const webUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `${webUrl}/sites/apps`;
  const solutionId = 'ac555cb1-e5ac-409e-86dc-61e6651b1e66';
  const clientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldCommandSet\",\"componentType\":\"Extension\",\"extensionType\":\"ListViewCommandSet\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-command-set\",\"scriptResources\":{\"hello-world-command-set\":{\"type\":\"path\",\"path\":\"hello-world-command-set_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldCommandSetStrings\":{\"type\":\"path\",\"path\":\"HelloWorldCommandSetStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
  const solution = { "FileSystemObjectType": 0, "Id": 40, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "ClientComponentId": clientSideComponentId, "ClientComponentManifest": clientComponentManifest, "SolutionId": solutionId, "Created": "2022-11-03T11:25:17", "Modified": "2022-11-03T11:26:03" };
  const solutionResponse = [solution];
  const application = { "FileSystemObjectType": 0, "Id": 31, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "SkipFeatureDeployment": true, "ContainsTenantWideExtension": true, "Modified": "2022-11-03T11:26:03", "CheckoutUserId": null, "EditorId": 9 };
  const applicationResponse = [application];

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommand,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_COMMANDSET_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds a tenant-wide ListView Command Set for lists', async () => {
    let executeCommandCalled = false;
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify(applicationResponse) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemAddCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, listType: 'List', title: customizerTitle, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('adds a tenant-wide ListView Command Set for libraries', async () => {
    let executeCommandCalled = false;
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify(applicationResponse) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemAddCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, listType: 'Library', location: 'ContextMenu', title: customizerTitle, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('adds a tenant-wide ListView Command Set for the SitePages library', async () => {
    let executeCommandCalled = false;
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify(applicationResponse) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemAddCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, listType: 'SitePages', location: 'CommandBar', title: customizerTitle, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('adds a tenant-wide ListView Command Set to a specific webtemplate and location including clientSideComponentProperties', async () => {
    let executeCommandCalled = false;
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify(solutionResponse) };
        }
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/AppCatalog`) {
          return { 'stdout': JSON.stringify(applicationResponse) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoListItemAddCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, listType: 'Library', title: customizerTitle, webTemplate: webTemplate, location: 'Both', clientSideComponentProperties: clientSideComponentProperties, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('throws an error when no app catalog is found', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': null };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError('Cannot add tenant-wide ListView Command Set as app catalog cannot be found'));
  });

  it('throws an error when specific client side component is not found in manifest list', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          return { 'stdout': JSON.stringify([]) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError('No component found with the specified clientSideComponentId found in the component manifest list. Make sure that the application is added to the application catalog'));
  });

  it('throws an error when the manifest of a specific client side component is not of type ListView Command Set', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoListItemListCommand) {
        if (args.options.listUrl === `${urlUtil.getServerRelativeSiteUrl(appCatalogUrl)}/Lists/ComponentManifests`) {
          const faultyClientComponentManifest = "{\"id\":\"6b2a54c5-3317-49eb-8621-1bbb76263629\",\"alias\":\"HelloWorldCommandSet\",\"componentType\":\"Extension\",\"extensionType\":\"FormCustomizer\",\"version\":\"0.0.1\",\"manifestVersion\":2,\"loaderConfig\":{\"internalModuleBaseUrls\":[\"HTTPS://SPCLIENTSIDEASSETLIBRARY/\"],\"entryModuleId\":\"hello-world-command-set\",\"scriptResources\":{\"hello-world-command-set\":{\"type\":\"path\",\"path\":\"hello-world-command-set_b47769f9eca3d3b6c4d5.js\"},\"HelloWorldCommandSetStrings\":{\"type\":\"path\",\"path\":\"HelloWorldCommandSetStrings_en-us_72ca11838ac9bae2790a8692c260e1ac.js\"},\"@microsoft/sp-application-base\":{\"type\":\"component\",\"id\":\"4df9bb86-ab0a-4aab-ab5f-48bf167048fb\",\"version\":\"1.15.2\"},\"@microsoft/sp-core-library\":{\"type\":\"component\",\"id\":\"7263c7d0-1d6a-45ec-8d85-d4d1d234171b\",\"version\":\"1.15.2\"}}},\"mpnId\":\"Undefined-1.15.2\",\"clientComponentDeveloper\":\"\"}";
          const solutionDuplicate = { ...solution };
          solutionDuplicate.ClientComponentManifest = faultyClientComponentManifest;
          return { 'stdout': JSON.stringify([solutionDuplicate]) };
        }
      }
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
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
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
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
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
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
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return { 'stdout': appCatalogUrl };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } }),
      new CommandError(`The solution has not been deployed to all sites. Make sure to deploy this solution to all sites.`));
  });

  it('fails validation if clientSideComponentId is not a valid Guid', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, listType: 'List', clientSideComponentId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if location value is not valid', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, listType: 'List', clientSideComponentId: clientSideComponentId, location: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if listType value is not valid', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, listType: 'invalid', clientSideComponentId: clientSideComponentId, location: 'Both' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all properties are specified', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, listType: 'List', location: 'Both', webTemplate: webTemplate, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});