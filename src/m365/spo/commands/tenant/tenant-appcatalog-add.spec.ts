import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./tenant-appcatalog-add');
import * as spoTenantAppCatalogUrlGetCommand from './tenant-appcatalogurl-get';
import * as spoSiteGetCommand from '../site/site-get';
import * as spoSiteRemoveCommand from '../site/site-remove';
import * as spoSiteClassicAddCommand from '../site/site-classic-add';
import * as assert from 'assert';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli';

describe(commands.TENANT_APPCATALOG_ADD, () => {
  let log: any[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
  });

  afterEach(() => {
    Utils.restore([
      Cli.executeCommand,
      Cli.executeCommandWithOutput
    ]);
  });

  after(() => {
    Utils.restore([
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

  it('creates app catalog when app catalog and site with different URL already exist and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used (debug)', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog and site with different URL already exist and force used failed', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when app catalog and site with different URL already exist, force used and deleting the existing site failed', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('Error deleting site new-app-catalog'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        return Promise.reject('Should not be called');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'Error deleting site new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used (debug)', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog already exists, site with different URL does not exist and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving site with different URL failed and app catalog already exists, and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when deleting existing app catalog failed', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error app catalog exists and no force used', (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/old-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when deleting existing site, when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when app catalog does not exist, site with different URL already exists and force not used', (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog and site with different URL don't exist`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog fails, when app catalog when app catalog does and site with different URL don't exist`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when checking if the app catalog site exists`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used (debug)`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when deleting existing site when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteRemoveCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when app catalog not registered, site with different URL exists and force not used`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
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

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered and site with different URL doesn't exist`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog when app catalog not registered and site with different URL doesn't exist`, (done) => {
    sinon.stub(Cli, 'executeCommand').callsFake((commandName, command, args) => {
      if (command === spoSiteClassicAddCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when app catalog not registered and checking if the site with different URL exists throws error`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (args.options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when checking if app catalog registered throws error`, (done) => {
    sinon.stub(Cli, 'executeCommandWithOutput').callsFake((commandName, command, args): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.reject(new CommandError('An error has occurred'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.strictEqual(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if the specified url is not a valid SharePoint URL', () => {
    const options: any = { url: '/foo', owner: 'user@contoso.com', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(typeof actual, 'string');
  });

  it('fails validation if timeZone is not a number', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 'a' };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(typeof actual, 'string');
  });

  it('passes validation when all options are specified and valid', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, true);
  });
});