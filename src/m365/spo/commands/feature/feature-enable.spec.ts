import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./feature-enable');

describe(commands.FEATURE_ENABLE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    requests = [];
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
    assert.strictEqual(command.name.startsWith(commands.FEATURE_ENABLE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures scope as string option', () => {
    const types = command.types;
    ['s', 'scope'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('Enable web feature (scope not defined, so defaults to web), no force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, featureId: 'b2307a39-e878-458b-bc90-03bc578531d6', url: 'https://contoso.sharepoint.com' } }, () => {
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
        sinonUtil.restore(request.post);
      }
    });
  });

  it('Enable site feature, force', (done) => {
    const requestUrl = `https://contoso.sharepoint.com/_api/site/features/add(featureId=guid'915c240e-a6cc-49b8-8b2c-0bff8b553ed3',force=true)`;
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(requestUrl) > -1) {
        if (opts.headers &&
          opts.headers.accept &&
          (opts.headers.accept as string).indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, featureId: '915c240e-a6cc-49b8-8b2c-0bff8b553ed3', url: 'https://contoso.sharepoint.com', scope: 'site', force: true } }, () => {
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
        sinonUtil.restore(request.post);
      }
    });
  });

  it('correctly handles enable feature reject request', (done) => {
    const err = 'Invalid enable feature reject request';
    const requestUrl = `https://contoso.sharepoint.com/_api/web/features/add(featureId=guid'b2307a39-e878-458b-bc90-03bc578531d6',force=false)`;

    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(requestUrl) > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
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
    const options = command.options;
    let containsScopeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('[scope]') > -1) {
        containsScopeOption = true;
      }
    });
    assert(containsScopeOption);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({
      options:
      {
        url: 'foo'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the required options specified', async () => {
    const actual = await command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Site', async () => {
    const actual = await command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Site'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts scope to be Web', async () => {
    const actual = await command.validate({
      options:
      {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: 'Web'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid string scope', async () => {
    const scope = 'foo';
    const actual = await command.validate({
      options: {
        url: "https://contoso.sharepoint.com",
        featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
        scope: scope
      }
    }, commandInfo);
    assert.strictEqual(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('doesn\'t fail validation if the optional scope option not specified', async () => {
    const actual = await command.validate(
      {
        options:
        {
          featureId: "00bfea71-5932-4f9c-ad71-1557e5751100",
          url: "https://contoso.sharepoint.com"
        }
      }, commandInfo);
    assert.strictEqual(actual, true);
  });
});