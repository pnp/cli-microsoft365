import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./hidedefaultthemes-get');

describe(commands.HIDEDEFAULTTHEMES_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      request.get,
      request.post,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.HIDEDEFAULTTHEMES_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('uses correct API url', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        return Promise.resolve('Correct Url');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false
      }
    }, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post
        ]);
      }
    });
  });

  it('uses correct API url (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        return Promise.resolve('Correct Url');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore([
          request.post
        ]);
      }
    });
  });

  it('gets the current value of the HideDefaultThemes setting', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({ value: true });
        }
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, verbose: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(true), 'Invalid request');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.post);
      }
    });
  });

  it('gets the current value of the HideDefaultThemes setting - handle error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/thememanager/GetHideDefaultThemes') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, verbose: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        sinonUtil.restore(request.post);
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
});