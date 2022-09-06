import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError, CommandErrorWithOutput } from '../../../../Command';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
import * as spoSiteAddCommand from '../site/site-add';
import * as spoSiteGetCommand from '../site/site-get';
import * as spoSiteRemoveCommand from '../site/site-remove';
import * as spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
const command: Command = require('./tenant-appcatalog-add');

describe(commands.TENANT_APPCATALOG_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_APPCATALOG_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used (debug)', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve({
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        });
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it('handles error when creating app catalog when app catalog and site with different URL already exist and force used failed', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog and site with different URL already exist, force used and deleting the existing site failed', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('Error deleting site new-app-catalog');
        }

        return Promise.reject('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        return Promise.reject('Should not be called');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('Error deleting site new-app-catalog'));
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used (debug)', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it('handles error when creating app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving site with different URL failed and app catalog already exists, and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('An error has occurred')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing app catalog failed', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve({
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        });
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error app catalog exists and no force used', async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve({
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        });
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/old-app-catalog'));
  });

  it('creates app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it('handles error when creating app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing site, when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject('404 FILE NOT FOUND');
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog does not exist, site with different URL already exists and force not used', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake(() => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog and site with different URL don't exist`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any);
  });

  it(`handles error when creating app catalog fails, when app catalog when app catalog does and site with different URL don't exist`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if the app catalog site exists`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake(() => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve({
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        });
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any);
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used (debug)`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } } as any);
  });

  it(`handles error when creating app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when deleting existing site when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered, site with different URL exists and force not used`, async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any);
  });

  it(`handles error when creating app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(Cli, 'executeCommand').callsFake((command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND')));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered and checking if the site with different URL exists throws error`, async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject('An error has occurred');
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if app catalog registered throws error`, async () => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    await assert.rejects(command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } } as any), new CommandError('An error has occurred'));
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if the specified url is not a valid SharePoint URL', async () => {
    const options: any = { url: '/foo', owner: 'user@contoso.com', timeZone: 4 };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('fails validation if timeZone is not a number', async () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 'a' };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(typeof actual, 'string');
  });

  it('passes validation when all options are specified and valid', async () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 4 };
    const actual = await command.validate({ options: options }, commandInfo);
    assert.strictEqual(actual, true);
  });
});