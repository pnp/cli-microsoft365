import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./site-recyclebinitem-restore');

describe(commands.SITE_RECYCLEBINITEM_RESTORE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.SITE_RECYCLEBINITEM_RESTORE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('fails validation if the siteUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { siteUrl: 'foo', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ids option is not a valid GUID', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '9526' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the siteUrl and ids options are valid', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } });
    assert(actual);
  });

  it('passes validation if siteUrl and id are defined', () => {
    const actual = command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com', ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526' } });
    assert(actual);
  });

  it('restores all items from recycle', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/site/RecycleBin/RestoreByIds') > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    const result = command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526,1adcf0d6-3733-4c13-b883-c84a27905cfd'
      }
    }, () => {
      try {
        done();
      }
      catch (e) {
        done(e);
      }
    });

    assert.equal(result, undefined);
  });

  it('catch error when restores all items from recycle', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        output: 'json',
        debug: true,
        siteUrl: 'https://contoso.sharepoint.com',
        ids: '5fb84a1f-6ab5-4d07-a6aa-31bba6de9526,1adcf0d6-3733-4c13-b883-c84a27905cfd'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith('Could not restore items from recyclebin'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});