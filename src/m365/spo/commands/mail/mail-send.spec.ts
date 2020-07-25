import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./mail-send');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.MAIL_SEND, () => {
  let log: any[];
  let cmdInstance: any;
  let requests: any[];

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
      }
    };
    requests = [];
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
    assert.strictEqual(command.name.startsWith(commands.MAIL_SEND), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('Send an email to one recipient (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and from someone (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'someone@contoso.com', verbose: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and from someone', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'someone@contoso.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and from some peoples (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'user@contoso.com,someone@consotos.com', verbose: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and from some peoples', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', from: 'user@contoso.com,someone@consotos.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and CC someone (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', cc: 'someone@contoso.com', verbose: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and CC someone', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', cc: 'someone@contoso.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and BCC someone (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', bcc: 'someone@contoso.com', verbose: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient and BCC someone', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', bcc: 'someone@contoso.com' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient with additional header (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}', verbose: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('Send an email to one recipient with additional header', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/SP.Utilities.Utility.SendEmail`) > -1 &&
          r.body) {
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

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action({ options: { debug: false, webUrl: "https://contoso.sharepoint.com", to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email', additionalHeaders: '{"X-Custom": "My Custom Header Value"}' } }, (err?: any) => {
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

  it('supports specifying URL', () => {
    const options = (command.options() as CommandOption[]);
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if at least the webUrl \'to\', \'subject\' and \'body\' are sprecified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', to: 'user@contoso.com', subject: 'Subject of the email', body: 'Content of the email' } });
    assert.strictEqual(actual, true);
  });
});
