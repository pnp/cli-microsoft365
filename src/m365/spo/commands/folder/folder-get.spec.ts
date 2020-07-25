import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./folder-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FOLDER_GET, () => {
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
            return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" });
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
    assert.strictEqual(command.name.startsWith(commands.FOLDER_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should correctly handle folder get reject request', (done) => {
    stubGetResponses(new Promise((resolve, reject) => { reject('error1'); }));

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents',
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

  it('should show tip when folder get rejects with error code 500', (done) => {
    sinon.stub(request, 'get').rejects({ statusCode: 500 });

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents',
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('Please check the folder URL. Folder might not exist on the specified URL')));
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
        folderUrl: '/Shared Documents',
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "test1", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/test1", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should pass the correct url params to request', (done) => {
    const request = stubGetResponses();

    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com',
        folderUrl: '/Shared Documents',
      }
    }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/_api/web/GetFolderByServerRelativeUrl(\'%2FShared%20Documents\')');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should pass the correct url params to request (sites/test1)', (done) => {
    const request = stubGetResponses();

    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        folderUrl: 'Shared Documents/',
      }
    }, () => {
      try {
        const lastCall: any = request.lastCall.args[0];
        assert.strictEqual(lastCall.url, 'https://contoso.sharepoint.com/sites/test1/_api/web/GetFolderByServerRelativeUrl(\'%2Fsites%2Ftest1%2FShared%20Documents\')');
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', folderUrl: '/Shared Documents' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and folderUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', folderUrl: '/Shared Documents' } });
    assert.strictEqual(actual, true);
  });
});