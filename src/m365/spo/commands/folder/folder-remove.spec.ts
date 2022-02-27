import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./folder-remove');

describe(commands.FOLDER_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let requests: any[];
  let promptOptions: any;
  let stubPostResponses: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;

    stubPostResponses = (removeResp: any = null) => {
      return sinon.stub(request, 'post').callsFake((opts) => {
        if ((opts.url as string).indexOf('GetFolderByServerRelativeUrl') > -1) {
          if (removeResp) {
            return removeResp;
          }
          else {
            return Promise.resolve();
          }
        }
  
        return Promise.reject('Invalid request');
      });
    };
  });

  beforeEach(() => {
    log = [];
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
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
    assert.strictEqual(command.name.startsWith(commands.FOLDER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing folder when confirmation argument not passed', (done) => {
    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents' } }, () => {
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

  it('aborts removing folder when prompt not confirmed', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents' } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the folder when prompt confirmed', (done) => {
    stubPostResponses();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: 
      { debug: false, 
        webUrl: 'https://contoso.sharepoint.com', 
        folderUrl: '/Shared Documents/Folder1' 
      } }, () => {
      try {
        assert(loggerLogSpy.notCalled === true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send params for remove request', (done) => {
    const request: sinon.SinonStub = stubPostResponses();

    command.action(logger, { options: 
      { verbose: true, 
        webUrl: 'https://contoso.sharepoint.com', 
        folderUrl: '/Shared Documents/Folder1',
        confirm: true
      } }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents%2FFolder1\')');
        assert.strictEqual(lastCall.method, 'POST');
        assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send params for remove request for sites/test1', (done) => {
    const request: sinon.SinonStub = stubPostResponses();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: 
      { verbose: true, 
        webUrl: 'https://contoso.sharepoint.com/sites/test1', 
        folderUrl: '/Shared Documents/Folder1' 
      } }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/test1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2Ftest1%2FShared%20Documents%2FFolder1\')');
        assert.strictEqual(lastCall.method, 'POST');
        assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send params for recycle request when recycle is set to true', (done) => {
    const request: sinon.SinonStub = stubPostResponses();

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: 
      { 
        debug: true,
        webUrl: 'https://contoso.sharepoint.com', 
        folderUrl: '/Shared Documents/Folder1', 
        recycle: true 
      } }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents%2FFolder1\')/recycle()');
        assert.strictEqual(lastCall.method, 'POST');
        assert.strictEqual(lastCall.headers['X-HTTP-Method'], 'DELETE');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should show error on request reject', (done) => {
    stubPostResponses(new Promise((resp, rej) => rej('error1')));

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: 
      { 
        debug: true,
        webUrl: 'https://contoso.sharepoint.com', 
        folderUrl: '/Shared Documents/Folder1', 
        recycle: true 
      } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
  });

  it('supports specifying URL', () => {
    const options = command.options();
    let containsTypeOption = false;
    options.forEach(o => {
      if (o.option.indexOf('<webUrl>') > -1) {
        containsTypeOption = true;
      }
    });
    assert(containsTypeOption);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = command.validate({ options: { webUrl: 'foo', folderUrl: '/Shared Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and folderUrl specified', () => {
    const actual = command.validate({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents' } });
    assert.strictEqual(actual, true);
  });
});