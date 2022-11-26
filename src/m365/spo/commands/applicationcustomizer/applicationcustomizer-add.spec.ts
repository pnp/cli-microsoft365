import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import * as spoTenantAppCatalogUrlGetCommand from '../tenant/tenant-appcatalogurl-get';
import * as spoListItemListCommand from '../listitem/listitem-list';
import * as spoListItemAddCommand from '../listitem/listitem-add';
import * as spoCustomActionAddCommand from '../customaction/customaction-add';
import { urlUtil } from '../../../../utils/urlUtil';
const command: Command = require('./applicationcustomizer-add');

describe(commands.APPLICATIONCUSTOMIZER_ADD, () => {
  const clientSideComponentId = '9748c81b-d72e-4048-886a-e98649543743';
  const customizerTitle = 'Some Customizer';
  const webTemplate = 'GROUP#0';
  const webUrl = 'https://contoso.sharepoint.com';
  const appCatalogUrl = `${webUrl}/sites/apps`;
  const solutionId = 'ac555cb1-e5ac-409e-86dc-61e6651b1e66';
  const solution = { "FileSystemObjectType": 0, "Id": 40, "ServerRedirectedEmbedUri": null, "ServerRedirectedEmbedUrl": "", "ClientComponentId": clientSideComponentId, "SolutionId": solutionId, "Created": "2022-11-03T11:25:17", "Modified": "2022-11-03T11:26:03" };
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
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APPLICATIONCUSTOMIZER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds an application customizer to a specific site', async () => {
    let executeCommandCalled = false;
    sinon.stub(Cli, 'executeCommand').callsFake(async (command): Promise<void> => {
      if (command === spoCustomActionAddCommand) {
        executeCommandCalled = true;
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: customizerTitle, clientSideComponentId: clientSideComponentId, verbose: true } });
    assert.strictEqual(executeCommandCalled, true);
  });

  it('adds a tenant-wide application customizer', async () => {
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

    await command.action(logger, { options: { clientSideComponentId: clientSideComponentId, title: customizerTitle, verbose: true } });
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
      new CommandError('Cannot add tenant-wide application customizer as app catalog cannot be found'));
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

  it('throws an error when solution is not dpeloyed globally', async () => {
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
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint url', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, webUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both webUrl and webTemplate is passed', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, webUrl: webUrl, webTemplate: webTemplate } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation for a web scoped application customizer', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, webUrl: webUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation for a tenant wide application customizer', async () => {
    const actual = await command.validate({ options: { title: customizerTitle, clientSideComponentId: clientSideComponentId, webTemplate: webTemplate } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });
});