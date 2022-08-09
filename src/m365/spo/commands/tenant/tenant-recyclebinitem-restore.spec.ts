import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./tenant-recyclebinitem-restore');

describe(commands.TENANT_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    (command as any).currentContext = undefined;
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TENANT_RECYCLEBINITEM_RESTORE), true);
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

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, commandInfo);
    assert(actual);
  });

  it(`restores deleted site collection from the tenant recycle bin, without waiting for completion`, (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers &&
          JSON.stringify(opts.data) === JSON.stringify({
            siteUrl: 'https://contoso.sharepoint.com/sites/hr'
          })) {
          return Promise.resolve("{\"HasTimedout\":false,\"IsComplete\":true,\"PollingInterval\":15000}");
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when the site to restore is not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_api/SPOInternalUseOnly.Tenant/RestoreDeletedSite`) > -1) {
        if (opts.headers &&
          JSON.stringify(opts.data) === JSON.stringify({
            siteUrl: 'https://contoso.sharepoint.com/sites/hr'
          })) {
          return Promise.reject("{\"odata.error\":{\"code\":\"-2147024809, System.ArgumentException\",\"message\":{\"lang\":\"en-US\",\"value\":\"Unable to find the deleted site: https://contoso.sharepoint.com/sites/hr.\"}}}");
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { url: 'https://contoso.sharepoint.com/sites/hr' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("{\"odata.error\":{\"code\":\"-2147024809, System.ArgumentException\",\"message\":{\"lang\":\"en-US\",\"value\":\"Unable to find the deleted site: https://contoso.sharepoint.com/sites/hr.\"}}}")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});