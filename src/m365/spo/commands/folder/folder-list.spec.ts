import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./folder-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FOLDER_LIST, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let stubGetResponses: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;

    stubGetResponses = (getResp: any = null) => {
      return sinon.stub(request, 'get').callsFake((opts) => {
        if ((opts.url as string).indexOf('GetFolderByServerRelativeUrl') > -1) {
          if (getResp) {
            return getResp;
          } else {
            return Promise.resolve({value:[{"Exists":true,"IsWOPIEnabled":false,"ItemCount":2,"Name":"Test","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/Test","TimeCreated":"2018-04-23T21:29:40Z","TimeLastModified":"2018-04-23T21:32:13Z","UniqueId":"3e735407-9c9f-418b-8378-450a9888d815","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"velin12","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/velin12","TimeCreated":"2018-05-02T22:28:50Z","TimeLastModified":"2018-05-02T22:36:14Z","UniqueId":"edeb37c6-8502-4a35-9fa2-6934bfc30214","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"test111","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/test111","TimeCreated":"2018-05-02T23:21:45Z","TimeLastModified":"2018-05-02T23:21:45Z","UniqueId":"0ac3da45-cacf-4c31-9b38-9ef3697d5a66","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"Forms","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/Forms","TimeCreated":"2018-02-15T13:57:52Z","TimeLastModified":"2018-02-15T13:57:52Z","UniqueId":"cbb96da6-c2d8-4af0-9451-d534d5949371","WelcomePage":""}]});
          }
        }
  
        return Promise.reject('Invalid request');
      });
    }
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.FOLDER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should correctly handle folder get reject request', (done) => {
    stubGetResponses(new Promise((res, rej)=>rej('error1')));

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('error1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle folder get success request', (done) => {
    stubGetResponses();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith([{"Name":"Test","ServerRelativeUrl":"/sites/abc/Shared Documents/Test"},{"Name":"velin12","ServerRelativeUrl":"/sites/abc/Shared Documents/velin12"},{"Name":"test111","ServerRelativeUrl":"/sites/abc/Shared Documents/test111"},{"Name":"Forms","ServerRelativeUrl":"/sites/abc/Shared Documents/Forms"}]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all information for output type json', (done) => {
    stubGetResponses();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        output: 'json'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith([{"Exists":true,"IsWOPIEnabled":false,"ItemCount":2,"Name":"Test","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/Test","TimeCreated":"2018-04-23T21:29:40Z","TimeLastModified":"2018-04-23T21:32:13Z","UniqueId":"3e735407-9c9f-418b-8378-450a9888d815","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"velin12","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/velin12","TimeCreated":"2018-05-02T22:28:50Z","TimeLastModified":"2018-05-02T22:36:14Z","UniqueId":"edeb37c6-8502-4a35-9fa2-6934bfc30214","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"test111","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/test111","TimeCreated":"2018-05-02T23:21:45Z","TimeLastModified":"2018-05-02T23:21:45Z","UniqueId":"0ac3da45-cacf-4c31-9b38-9ef3697d5a66","WelcomePage":""},{"Exists":true,"IsWOPIEnabled":false,"ItemCount":0,"Name":"Forms","ProgID":null,"ServerRelativeUrl":"/sites/abc/Shared Documents/Forms","TimeCreated":"2018-02-15T13:57:52Z","TimeLastModified":"2018-02-15T13:57:52Z","UniqueId":"cbb96da6-c2d8-4af0-9451-d534d5949371","WelcomePage":""}]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send correct request params when /', (done) => {
    let request: sinon.SinonStub = stubGetResponses();

    cmdInstance.action({
      options: {
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
      }
    }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents\')/folders');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should send correct request params when /sites/abc', (done) => {
    let request: sinon.SinonStub = stubGetResponses();

    cmdInstance.action({
      options: {
        verbose: true,
        webUrl: 'https://contoso.sharepoint.com/sites/abc',
        parentFolderUrl: '/Shared Documents',
      }
    }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/abc/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2Fabc%2FShared%20Documents\')/folders');
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', parentFolderUrl: '/Shared Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and parentFolderUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', parentFolderUrl: '/Shared Documents' } });
    assert.strictEqual(actual, true);
  });
});