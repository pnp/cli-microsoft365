import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./listitem-roleinheritance-reset');

describe(commands.LISTITEM_ROLEINHERITANCE_RESET, () => {
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.LISTITEM_ROLEINHERITANCE_RESET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['listId', 'listTitle']]);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', listItemId: 8, listTitle: 'Demo List' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listItemId: 8, listId: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the listId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listItemId: 8, listId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the listId option is a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listItemId: 8, listId: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the specified list item id is not a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: 'a' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the specified list item id is a number', async () => {
    const actual = await command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', listTitle: 'Demo List', listItemId: '4' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('reset role inheritance on list item by list title', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(8)/resetroleinheritance') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listTitle: 'test'
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('reset role inheritance on list item by list id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid\'0cd891ef-afce-4e55-b836-fce03286cccf\')/items(8)/resetroleinheritance') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listId: '0cd891ef-afce-4e55-b836-fce03286cccf'
      }
    }, (err: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when reseting list item role inheritance', (done) => {
    const err = 'request rejected';
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/getbytitle(\'test\')/items(8)/resetroleinheritance') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        listItemId: 8,
        listTitle: 'test'
      }
    }, (error?: any) => {
      try {
        assert.strictEqual(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});