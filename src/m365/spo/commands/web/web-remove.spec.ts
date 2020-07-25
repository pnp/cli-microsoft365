import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./web-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.WEB_REMOVE, () => {
  let log: any[];
  let requests: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    promptOptions = undefined;
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
    assert.strictEqual(command.name.startsWith(commands.WEB_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('should fail validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'foo'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required options are specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('should prompt before deleting subsite when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { webUrl: 'https://contoso.sharepoint.com/subsite' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes web successfully without prompting with confirmation argument', (done) => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web`) > -1 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['accept'] === 'application/json;odata=nometadata') {
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
    });

  });

  it('deletes web successfully when prompt confirmed', (done) => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite"
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web`) > -1 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['accept'] === 'application/json;odata=nometadata') {
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
    });
  });

  it('deletes web successfully without prompting with confirmation argument (verbose)', (done) => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web`) > -1 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['accept'] === 'application/json;odata=nometadata') {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        assert(cmdInstanceLogSpy.calledWith(sinon.match('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('deletes web successfully without prompting with confirmation argument (debug)', (done) => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web`) > -1 &&
          r.headers['X-HTTP-Method'] === 'DELETE' &&
          r.headers['accept'] === 'application/json;odata=nometadata') {
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
    });
  });

  it('handles error when deleting web', (done) => {
    // Delete web
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});