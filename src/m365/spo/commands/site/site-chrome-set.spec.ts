import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./site-chrome-set');

describe(commands.SITE_CHROME_SET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.SITE_CHROME_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`doesn't return error on a valid request`, (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/web/SetChromeOptions`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`sends a request without any data when no options were specified`, (done) => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/test' } }, () => {
      try {
        assert.strictEqual(Object.keys(data).length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables mega menu', (done) => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/test', disableMegaMenu: "true" } }, () => {
      try {
        assert.strictEqual(data.megaMenuEnabled, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables footer', (done) => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/test', disableFooter: "true" } }, () => {
      try {
        assert.strictEqual(data.footerEnabled, false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables title in the header', (done) => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/test', hideTitleInHeader: "true" } }, () => {
      try {
        assert.strictEqual(data.hideTitleInHeader, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('configures chrome with enum values', (done) => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/test', headerLayout: "Extended", headerEmphasis: "Light", logoAlignment: "Center", footerLayout: "Extended", footerEmphasis: "Light" } }, () => {
      try {
        assert.strictEqual(data.headerLayout, 4);
        assert.strictEqual(data.headerEmphasis, 1);
        assert.strictEqual(data.logoAlignment, 1);
        assert.strictEqual(data.footerLayout, 2);
        assert.strictEqual(data.footerEmphasis, 2);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when setting site chrome settings', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    command.action(logger, { options: { debug: false, url: 'https://contoso.sharepoint.com/sites/Sales', footerEmphasis: 'Light' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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
      if (o.option.indexOf('<url>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying disable footer', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--disableFooter') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying disable mega menu', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--disableMegaMenu') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying hide title in header', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--hideTitleInHeader') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying footer emphasis', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[footerEmphasis]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the footer emphasis option is not a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', footerEmphasis: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the footer emphasis option is a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', footerEmphasis: "Dark" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying footer layout', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[footerLayout]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the footer layout option is not a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', footerLayout: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the footer layout option is a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', footerLayout: "Simple" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying header emphasis', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[headerEmphasis]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the header emphasis option is not a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', headerEmphasis: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the header emphasis option is a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', headerEmphasis: "Dark" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying header layout', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[headerLayout]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the header emphasis layout is not a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', headerLayout: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the header emphasis layout is a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', headerLayout: "Standard" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('supports specifying logo alignment', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[logoAlignment]') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the header logo alignment is not a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', logoAlignment: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the logo alignment is a valid', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com', logoAlignment: "Center" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});