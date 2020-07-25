import commands from '../../commands';
import sinon = require('sinon');
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import Utils from '../../../../Utils';
import request from '../../../../request';
import * as assert from 'assert';
import Command, { CommandValidate, CommandOption, CommandTypes, CommandError } from '../../../../Command';
const command: Command = require('./feature-disable');

describe(commands.FEATURE_DISABLE, () => {
  let log: string[];
  let cmdInstance: any;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    requests = [];
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
    assert.strictEqual(command.name.startsWith(commands.FEATURE_DISABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures scope as string option', () => {
    const types = (command.types() as CommandTypes);
    ['s', 'scope'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('fails validation if scope is not site|web', () => {
    const scope = 'list';
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        featureId: "780ac353-eaf8-4ac2-8c47-536d93c03fd6",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('passes validation if url and featureId is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        featureId: "780ac353-eaf8-4ac2-8c47-536d93c03fd6"
      }
    });

    assert.strictEqual(actual, true);

  });

  it('supports specifying scope', () => {
    const options = (command.options() as CommandOption[]);
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('disables web feature (scope not defined, so defaults to web), no force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=false)`
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, featureId: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', url: 'https://contoso.sharepoint.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(requestUrl) > -1 && r.headers.accept && r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('disables web feature (scope not defined, so defaults to web), with force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=true)`
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, featureId: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', url: 'https://contoso.sharepoint.com', force: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(requestUrl) > -1 && r.headers.accept && r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('disables site feature (scope explicitly set), no force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/site/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=false)`
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, featureId: '780ac353-eaf8-4ac2-8c47-536d93c03fd6', url: 'https://contoso.sharepoint.com', scope: 'site' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(requestUrl) > -1 && r.headers.accept && r.headers.accept.indexOf('application/json') === 0) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.post);
      }
    });
  });

  it('correctly handles disable feature reject request', (done) => {
    const err = 'Invalid disable feature reject request';
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/remove(featureId=guid'780ac353-eaf8-4ac2-8c47-536d93c03fd6',force=false)`

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(requestUrl) > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        url: 'https://contoso.sharepoint.com',
        featureId: "780ac353-eaf8-4ac2-8c47-536d93c03fd6",
        scope: 'web'
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