import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./user-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.USER_REMOVE, () => {
  let log: any[];
  let requests: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    assert.strictEqual(command.name.startsWith(commands.USER_REMOVE), true);
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

  it('fails validation if id or loginName options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails valiation if id or loginname oprions are passed', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        id: 10,
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    });
    assert.notStrictEqual(actual, true);
  })

  it('should fail validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({
      options:
      {
        webUrl: 'foo',
        id: 10
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('should prompt before removing user using id from web when confirmation argument not passed ', (done) => {
    cmdInstance.action({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        id: 10
      }
    }, () => {
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

  it('should prompt before removing user using login name from web when confirmation argument not passed ', (done) => {
    cmdInstance.action({
      options:
      {
        webUrl: 'https://contoso.sharepoint.com/subsite',
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    }, () => {
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

  it('removes user by id successfully without prompting with confirmation argument', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
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

  it('removes user by login name successfully without prompting with confirmation argument', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cjohn.doe%40mytenant.onmicrosoft.com')`) > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        loginName: "i:0#.f|membership|parker@tenant.onmicrosoft.com",
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removeByLoginName('i%3A0%23.f%7Cmembership%7Cparker%40tenant.onmicrosoft.com')`) > -1 && 
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

  it('removes user by id successfully from web when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
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

  it('removes user by login name successfully from web when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`_api/web/siteusers/removeByLoginName`) > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        loginName: "i:0#.f|membership|john.doe@mytenant.onmicrosoft.com"
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removeByLoginName`) > -1 &&
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

  it('removes user from web successfully without prompting with confirmation argument (verbose)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
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

  it('removes user from web successfully without prompting with confirmation argument (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return Promise.resolve(true);
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
        confirm: true
      }
    }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`_api/web/siteusers/removebyid(10)`) > -1 &&
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

  it('handles error when removing using from web', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf('_api/web/siteusers/removebyid(10)') > -1) {
        return Promise.reject('An error has occurred');
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        webUrl: "https://contoso.sharepoint.com/subsite",
        id: 10,
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