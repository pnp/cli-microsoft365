import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./file-remove');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.FILE_REMOVE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.post
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FILE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.FILE_REMOVE);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing file when confirmation argument not passed (id)', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
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

  it('prompts before removing file when confirmation argument not passed (title)', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
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

  it('aborts removing file when prompt not confirmed', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the file when prompt confirmed (id)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/web/GetFileById(guid'`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/GetFileById(guid'`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl does not includes a trailing /', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      console.log(opts.url);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl includes a trailing /', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl does not includes a trailing / and fileUrl is server relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com';
    const fileUrl: string = '/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl includes a trailing / and fileUrl is server relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/';
    const fileUrl: string = '/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) does not includes a trailing / ', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) includes a trailing /', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) does not includes a trailing / and fileUrl is server relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = '/sites/subsite/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) includes a trailing / and fileUrl is server relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = '/sites/subsite/SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent(fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) does not includes a trailing / and fileUrl is site relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when webUrl (subsite) includes a trailing / and fileUrl is site relative', (done) => {
    const siteUrl: string = 'https://contoso.sharepoint.com/sites/subsite/';
    const fileUrl: string = 'SharedDocuments/Document.docx';

    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: siteUrl, url: fileUrl } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`GetFileByServerRelativeUrl('${encodeURIComponent('/sites/subsite/'+fileUrl)}')`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('recycles the file when prompt confirmed (id)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/recycle()`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', id: '0cd891ef-afce-4e55-b836-fce03286cccf', recycle: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/recycle()`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('removes the file when prompt confirmed (url)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/_api/web/GetFileByServerRelativeUrl('`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', url: '0cd891ef-afce-4e55-b836-fce03286cccf' } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/web/GetFileByServerRelativeUrl('`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('recycles the file when prompt confirmed (url)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if (opts.url.indexOf(`/recycle()`) > -1) {
        if (opts.headers.authorization &&
          opts.headers.authorization.indexOf('Bearer ') === 0 &&
          opts.headers.accept &&
          opts.headers.accept.indexOf('application/json') === 0) {
          return Promise.resolve();
        }
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', url: '0cd891ef-afce-4e55-b836-fce03286cccf', recycle: true } }, () => {
      let correctRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/recycle()`) > -1 &&
          r.headers.authorization &&
          r.headers.authorization.indexOf('Bearer ') === 0 &&
          r.headers.accept &&
          r.headers.accept.indexOf('application/json') === 0) {
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

  it('command correctly handles file remove reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFileById(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = '0cd891ef-afce-4e55-b836-fce03286cccf';

    cmdInstance.action({
      options: {
        debug: true,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {

      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });
  });

  it('uses correct API url when id option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFileById(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });

  });

  it('uses correct API url when url option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFileByServerRelativeUrl(') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionUrl: string = 'SharedDocuments/Test.docx';

    cmdInstance.action({
      options: {
        debug: false,
        url: actionUrl,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
      }
    });

  });

  it('uses correct API url when recycle option is passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/recycle()') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    const actionId: string = '0cd891ef-afce-4e55-b836-fce03286cccf';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,        
        recycle: true,
        webUrl: 'https://contoso.sharepoint.com',
        confirm: true
      }
    }, () => {

      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post
        ]);
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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if both id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', id: '0cd891ef-afce-4e55-b836-fce03286cccf' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if the id option is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '12345' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the id option is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF' } });
    assert(actual);
  });

  it('fails validation if both id and url options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', url: 'Documents' } });
    assert.notEqual(actual, true);
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
    assert(find.calledWith(commands.FILE_REMOVE));
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

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.getAccessToken);
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        id: "BC448D63-484F-49C5-AB8C-96B14AA68D50",
        webUrl: "https://contoso.sharepoint.com",
        debug: false,
        confirm: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});