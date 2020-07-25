import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./theme-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.THEME_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let promptOptions: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
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
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.THEME_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should prompt before removing theme when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { debug: false, name: 'Contoso' } }, () => {
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

  it('removes theme successfully without prompting with confirmation argument', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso',
        confirm: true
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
        assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
        assert.strictEqual(postStub.lastCall.args[0].body.name, 'Contoso');
        assert.strictEqual(postStub.lastCall.args[0].json, true);
        assert.strictEqual(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes theme successfully without prompting with confirmation argument (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        confirm: true
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
        assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
        assert.strictEqual(postStub.lastCall.args[0].body.name, 'Contoso');
        assert.strictEqual(postStub.lastCall.args[0].json, true);
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes theme successfully when prompt confirmed', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_api/thememanager/DeleteTenantTheme');
        assert.strictEqual(postStub.lastCall.args[0].headers['accept'], 'application/json;odata=nometadata');
        assert.strictEqual(postStub.lastCall.args[0].body.name, 'Contoso');
        assert.strictEqual(postStub.lastCall.args[0].json, true);
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when removing theme', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {

      if ((opts.url as string).indexOf('/_api/thememanager/DeleteTenantTheme') > -1) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
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
});