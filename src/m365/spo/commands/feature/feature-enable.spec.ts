import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError, CommandTypes } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./feature-enable');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FEATURE_ENABLE, () => {
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
    assert.strictEqual(command.name.startsWith(commands.FEATURE_ENABLE), true);
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

  it('Enable web feature (scope not defined, so defaults to web), no force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`
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

    cmdInstance.action({ options: { debug: true, featureId: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' } }, () => {
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

  it('Enable site feature, force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/site/features/add(featureId=guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3',force=true)`
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

    cmdInstance.action({ options: { debug: true, featureId: '915c240e-a6cc-49b8-8b2c-0bff8b553ed3', url: 'https://contoso.sharepoint.com', scope: 'site', force: true } }, () => {
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

  it('correctly handles enable feature reject request', (done) => {
    const err = 'Invalid enable feature reject request';
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`

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
        featureId: "b2307a39-e878-458b-bc90-03bc578531d6",
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

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the required options specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Site', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Site'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Web'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid string scope', () => {
    const scope = 'foo';
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: scope
      }
    });
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });
  
  it('doesn\'t fail validation if the optional scope option not specified', () => {
    const actual = (command.validate() as CommandValidate)(
      {
        options:
        {
          featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
          url: "https://contoso.sharepoint.com"
        }
      });
    assert.strictEqual(actual, true);
  });
});