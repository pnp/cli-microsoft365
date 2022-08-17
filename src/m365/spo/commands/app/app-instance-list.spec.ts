import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./app-instance-list');

describe(commands.APP_INSTANCE_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.APP_INSTANCE_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), [`Title`, `AppId`]);
  });

  it('fails validation when siteUrl is not a valid url', async () => {
    const actual = await command.validate({ options: { siteUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the siteUrl is a valid url', async () => {
    const actual = await command.validate({ options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves available apps from the site collection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/AppTiles') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve({
            value: [
              {
                AppId: 'b2307a39-e878-458b-bc90-03bc578531d6',
                Title: 'online-client-side-solution'
              },
              {
                AppId: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
                Title: 'onprem-client-side-solution'
              }
            ]
          });
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite' } }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            AppId: 'b2307a39-e878-458b-bc90-03bc578531d6',
            Title: 'online-client-side-solution'
          },
          {
            AppId: 'e5f65aef-68fe-45b0-801e-92733dd57e2c',
            Title: 'onprem-client-side-solution'
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });



  it('correctly handles no apps found in the site collection', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/AppTiles') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite', debug: false } }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles no apps found in the site collection (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/AppTiles') > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve(JSON.stringify({ value: [] }));
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { siteUrl: 'https://contoso.sharepoint.com/sites/testsite', verbose: true } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith('No apps found'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error while listing apps in the site collection', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');

    });

    command.action(logger, {
      options: {
        siteUrl: 'https://contoso.sharepoint.com/sites/testsite'
      }
    } as any, (err?: any) => {
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
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });
});