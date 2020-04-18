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
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.FEATURE_DISABLE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('configures command types', () => {
    assert.notEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures scope as string option', () => {
    const types = (command.types() as CommandTypes);
    ['s', 'scope'].forEach(o => {
      assert.notEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('fails validation if the url is not provided', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        force: false,
        scope: 'web',
        featureId: '780ac353-eaf8-4ac2-8c47-536d93c03fd6'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url is empty', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: '',
        force: false,
        scope: 'web',
        featureId: '780ac353-eaf8-4ac2-8c47-536d93c03fd6'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the featureId is not provided', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/sales',
        force: false,
        scope: 'web'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the featureId is empty', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: 'https://contoso.sharepoint.com/sites/sales',
        force: false,
        scope: 'web',
        featureId: ''
      }
    });
    assert.notEqual(actual, true);
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
    assert.equal(actual, `${scope} is not a valid Feature scope. Allowed values are Site|Web`);
  });

  it('passes validation if url and featureId is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        url: "https://contoso.sharepoint.com",
        featureId: "780ac353-eaf8-4ac2-8c47-536d93c03fd6"
      }
    });

    assert.equal(actual, true);

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
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.FEATURE_DISABLE));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('has help with remarks', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsRemarks: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Remarks:') > -1) {
        containsRemarks = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsRemarks);
  });
});