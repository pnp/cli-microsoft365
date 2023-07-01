import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-chrome-set.js';

describe(commands.SITE_CHROME_SET, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_CHROME_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it(`doesn't return error on a valid request`, async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/Sales/_api/web/SetChromeOptions`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales' } });
  });

  it(`sends a request without any data when no options were specified`, async () => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/test' } });
    assert.strictEqual(Object.keys(data).length, 0);
  });

  it('disables mega menu', async () => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/test', disableMegaMenu: true } });
    assert.strictEqual(data.megaMenuEnabled, false);
  });

  it('disables footer', async () => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/test', disableFooter: true } });
    assert.strictEqual(data.footerEnabled, false);
  });

  it('disables title in the header', async () => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/test', hideTitleInHeader: true } });
    assert.strictEqual(data.hideTitleInHeader, true);
  });

  it('configures chrome with enum values', async () => {
    let data: any = {};
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso.sharepoint.com/sites/test/_api/web/SetChromeOptions`) {
        data = opts.data;
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/test', headerLayout: "Extended", headerEmphasis: "Light", logoAlignment: "Center", footerLayout: "Extended", footerEmphasis: "Light" } });
    assert.strictEqual(data.headerLayout, 4);
    assert.strictEqual(data.headerEmphasis, 1);
    assert.strictEqual(data.logoAlignment, 1);
    assert.strictEqual(data.footerLayout, 2);
    assert.strictEqual(data.footerEmphasis, 2);
  });

  it('correctly handles OData error when setting site chrome settings', async () => {
    sinon.stub(request, 'post').rejects({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });

    await assert.rejects(command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/Sales', footerEmphasis: 'Light' } } as any), new CommandError('An error has occurred'));
  });

  it('supports specifying URL', () => {
    const options = command.options;
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<siteUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com' } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', footerEmphasis: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the footer emphasis option is a valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', footerEmphasis: "Dark" } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', footerLayout: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the footer layout option is a valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', footerLayout: "Simple" } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', headerEmphasis: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the header emphasis option is a valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', headerEmphasis: "Dark" } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', headerLayout: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the header emphasis layout is a valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', headerLayout: "Standard" } }, commandInfo);
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
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', logoAlignment: "None" } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the logo alignment is a valid', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', logoAlignment: "Center" } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
