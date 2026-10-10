import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError, CommandErrorWithOutput } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import spoSiteAddCommand from '../site/site-add.js';
import spoSiteGetCommand from '../site/site-get.js';
import spoSiteRemoveCommand from '../site/site-remove.js';
import command, { options } from './tenant-appcatalog-add.js';
import spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get.js';

describe(commands.TENANT_APPCATALOG_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      cli.executeCommand,
      cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_APPCATALOG_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used', async () => {
    const executeCommandStub = sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com',
      timeZone: '4',
      force: true 
    }) });

    const siteRemoveCalls = executeCommandStub.getCalls().filter(call => call.args[0] === spoSiteRemoveCommand);
    assert.strictEqual(siteRemoveCalls.length, 1);
    assert.strictEqual(siteRemoveCalls[0].args[1].options.permanent, true);
    assert.strictEqual(siteRemoveCalls[0].args[1].options.wait, true);
    assert.strictEqual(siteRemoveCalls[0].args[1].options.force, true);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used (debug)', async () => {
    const executeCommandStub = sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return {
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        };
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true,
        debug: true
      })
    });

    const siteRemoveCalls = executeCommandStub.getCalls().filter(call => call.args[0] === spoSiteRemoveCommand);
    assert.strictEqual(siteRemoveCalls.length, 2);
    siteRemoveCalls.forEach(call => {
      assert.strictEqual(call.args[1].options.permanent, true);
      assert.strictEqual(call.args[1].options.wait, true);
      assert.strictEqual(call.args[1].options.force, true);
    });
  });

  it('handles error when creating app catalog when app catalog and site with different URL already exist and force used failed', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    }), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog and site with different URL already exist, force used and deleting the existing site failed', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('Error deleting site new-app-catalog');
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        throw 'Should not be called';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    }), new CommandError('Error deleting site new-app-catalog'));
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    });
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used (debug)', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true, 
      debug: true 
    }) });
  });

  it('handles error when creating app catalog when app catalog already exists, site with different URL does not exist and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true, 
      debug: true 
    }) }), new CommandError('An error has occurred'));
  });

  it('handles error when retrieving site with different URL failed and app catalog already exists, and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('An error has occurred'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true, 
      debug: true 
    }) }), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing app catalog failed', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return {
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        };
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true 
    }) }), new CommandError('An error has occurred'));
  });

  it('handles error app catalog exists and no force used', async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return {
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        };
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4'
      })
    }), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/old-app-catalog'));
  });

  it('creates app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
        owner: 'user@contoso.com', 
        timeZone: '4', 
        force: true
      })
    });
  });

  it('handles error when creating app catalog when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true 
    }) }), new CommandError('An error has occurred'));
  });

  it('handles error when deleting existing site, when app catalog does not exist, site with different URL already exists and force used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4', 
      force: true 
    }) }), new CommandError('An error has occurred'));
  });

  it('handles error when app catalog does not exist, site with different URL already exists and force not used', async () => {
    sinon.stub(cli, 'executeCommand').callsFake(() => {
      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog', 
      owner: 'user@contoso.com', 
      timeZone: '4'
    }) }), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog and site with different URL don't exist`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
      owner: 'user@contoso.com',
      timeZone: '4'
    }) });
  });

  it(`handles error when creating app catalog fails, when app catalog when app catalog does and site with different URL don't exist`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return 'https://contoso.sharepoint.com/sites/old-app-catalog';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4'
      })
    }), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if the app catalog site exists`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(() => {
      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return {
          stdout: 'https://contoso.sharepoint.com/sites/old-app-catalog'
        };
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('An error has occurred'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4'
      })
    }), new CommandError('An error has occurred'));
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    });
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used (debug)`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true,
        debug: true
      })
    });
  });

  it(`handles error when creating app catalog when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    }), new CommandError('An error has occurred'));
  });

  it(`handles error when deleting existing site when app catalog not registered, site with different URL exists and force used`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4',
        force: true
      })
    }), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered, site with different URL exists and force not used`, async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
        owner: 'user@contoso.com',
        timeZone: '4'
      })
    }), new CommandError('Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog'));
  });

  it(`creates app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return;
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
      owner: 'user@contoso.com',
      timeZone: '4'
    }) });
  });

  it(`handles error when creating app catalog when app catalog not registered and site with different URL doesn't exist`, async () => {
    sinon.stub(cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === spoSiteAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandError('An error has occurred');
        }

        throw 'Invalid URL';
      }

      throw 'Unknown case';
    });
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('404 FILE NOT FOUND'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
      owner: 'user@contoso.com',
      timeZone: '4'
    }) }), new CommandError('An error has occurred'));
  });

  it(`handles error when app catalog not registered and checking if the site with different URL exists throws error`, async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake(async (command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return '';
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          throw new CommandErrorWithOutput(new CommandError('An error has occurred'));
        }

        throw new CommandError('Invalid URL');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
      owner: 'user@contoso.com',
      timeZone: '4'
    }) }), new CommandError('An error has occurred'));
  });

  it(`handles error when checking if app catalog registered throws error`, async () => {
    sinon.stub(cli, 'executeCommandWithOutput').callsFake((command): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        throw new CommandError('An error has occurred');
      }

      throw 'Unknown case';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ 
      url: 'https://contoso.sharepoint.com/sites/new-app-catalog',
      owner: 'user@contoso.com',
      timeZone: '4'
    }) }), new CommandError('An error has occurred'));
  });

  it('fails validation if the specified url is not a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'abc', owner: 'user@contoso.com', timeZone: '4' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if timeZone is not a number', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 'a' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when all options are specified and valid', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: '4' });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com/sites/apps', unknownOption: 'value' });
    assert.strictEqual(actual.success, false);
  });
});
