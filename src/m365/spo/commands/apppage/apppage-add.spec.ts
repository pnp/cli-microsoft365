import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./apppage-add');

describe(commands.APPPAGE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

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
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.APPPAGE_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a single-part app page', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.data.webPartDataAsJson ===
        "{}" && !opts.data.addToQuickLaunch) {
        return Promise.resolve("Done");
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } }, () => {
      try {
        assert(loggerSpy.calledWith("Done"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a single-part app page showing on quicklaunch', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.data.webPartDataAsJson ===
        "{}" && opts.data.addToQuickLaunch) {
        return Promise.resolve("Done");
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, addToQuickLaunch: true, title: 'test-single', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } }, () => {
      try {
        assert(loggerSpy.calledWith("Done"));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to create a single-part app page if request is rejected', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`_api/sitepages/Pages/CreateFullPageApp`) > -1 &&
        opts.data.title === "failme") {
        return Promise.reject('Failed to create a single-part app page');
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, title: 'failme', webUrl: 'https://contoso.sharepoint.com/', webPartData: JSON.stringify({}) } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Failed to create a single-part app page`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying title', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webUrl', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webPartData', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webPartData') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
  it('fails validation if webPartData is not a valid JSON string', () => {
    const actual = command.validate({ options: { title: 'Contoso', webUrl: 'https://contoso', webPartData: 'abc' } });
    assert.notStrictEqual(actual, true);
  });
  it('validation passes on all required options', () => {
    const actual = command.validate({ options: { title: 'Contoso', webPartData: '{}', webUrl: 'https://contoso.sharepoint.com' } });
    assert.strictEqual(actual, true);
  });
});