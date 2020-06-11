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

describe(commands.TENANT_APPCATALOG_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
      Utils.executeCommand,
      Utils.executeCommandWithOutput
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
    assert.equal(command.name.startsWith(commands.TENANT_APPCATALOG_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog and site with different URL already exist and force used (debug)', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog and site with different URL already exist and force used failed', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when app catalog and site with different URL already exist, force used and deleting the existing site failed', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('Error deleting site new-app-catalog'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        return Promise.reject('Should not be called');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'Error deleting site new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog already exists, site with different URL does not exist and force used (debug)', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog already exists, site with different URL does not exist and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when retrieving site with different URL failed and app catalog already exists, and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when deleting existing app catalog failed', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error app catalog exists and no force used', (done) => {
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/old-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates app catalog when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating app catalog when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when deleting existing site, when app catalog does not exist, site with different URL already exists and force used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when app catalog does not exist, site with different URL already exists and force not used', (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog and site with different URL don't exist`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog fails, when app catalog when app catalog does and site with different URL don't exist`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog' ||
          options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when checking if the app catalog site exists`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('https://contoso.sharepoint.com/sites/old-app-catalog');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/old-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered, site with different URL exists and force used (debug)`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true, debug: true } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when deleting existing site when app catalog not registered, site with different URL exists and force used`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteRemoveCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog', force: true } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when app catalog not registered, site with different URL exists and force not used`, (done) => {
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'Another site exists at https://contoso.sharepoint.com/sites/new-app-catalog');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`creates app catalog when app catalog not registered and site with different URL doesn't exist`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.resolve();
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when creating app catalog when app catalog not registered and site with different URL doesn't exist`, (done) => {
    sinon.stub(Utils, 'executeCommand').callsFake((command, options, cmd) => {
      if (command === spoSiteClassicAddCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject('Invalid URL');
      }

      return Promise.reject(new CommandError('Unknown case'));
    });
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('404 FILE NOT FOUND'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when app catalog not registered and checking if the site with different URL exists throws error`, (done) => {
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.resolve('');
      }

      if (command === spoSiteGetCommand) {
        if (options.url === 'https://contoso.sharepoint.com/sites/new-app-catalog') {
          return Promise.reject(new CommandError('An error has occurred'));
        }

        return Promise.reject(new CommandError('Invalid URL'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`handles error when checking if app catalog registered throws error`, (done) => {
    sinon.stub(Utils, 'executeCommandWithOutput').callsFake((command, options, cmd): Promise<any> => {
      if (command === spoTenantAppCatalogUrlGetCommand) {
        return Promise.reject(new CommandError('An error has occurred'));
      }

      return Promise.reject(new CommandError('Unknown case'));
    });

    cmdInstance.action({ options: { url: 'https://contoso.sharepoint.com/sites/new-app-catalog' } }, (err: CommandError) => {
      try {
        assert.equal(err.message, 'An error has occurred');
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

  it('fails validation if url not specified', () => {
    const options: any = { owner: 'user@contoso.com', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.equal(typeof actual, 'string');
  });

  it('fails validation if the specified url is not a valid SharePoint URL', () => {
    const options: any = { url: '/foo', owner: 'user@contoso.com', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.equal(typeof actual, 'string');
  });

  it('fails validation if owner not specified', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.equal(typeof actual, 'string');
  });

  it('fails validation if timeZone not specified', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com' };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.equal(typeof actual, 'string');
  });

  it('fails validation if timeZone is not a number', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 'a' };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.equal(typeof actual, 'string');
  });

  it('passes validation when all options are specified and valid', () => {
    const options: any = { url: 'https://contoso.sharepoint.com/sites/apps', owner: 'user@contoso.com', timeZone: 4 };
    const actual = (command.validate() as CommandValidate)({ options: options });
    assert.strictEqual(actual, true);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TENANT_APPCATALOG_ADD));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});