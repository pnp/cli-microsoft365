import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./folder-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FOLDER_ADD, () => {
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let stubPostResponses: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;

    stubPostResponses = (addResp: any = null) => {
      return sinon.stub(request, 'post').callsFake((opts) => {
        if ((opts.url as string).indexOf('/_api/web/folders') > -1) {
          if (addResp) {
            return addResp;
          } else {
            return Promise.resolve({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" });
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
    assert.strictEqual(command.name.startsWith(commands.FOLDER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('should correctly handle folder add reject request', (done) => {
    stubPostResponses(new Promise((resolve, reject) => { reject('error1'); }));

    cmdInstance.action({
      options: {
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
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

  it('should correctly handle folder add success request', (done) => {
    stubPostResponses();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.lastCall.calledWith({ "Exists": true, "IsWOPIEnabled": false, "ItemCount": 0, "Name": "abc", "ProgID": null, "ServerRelativeUrl": "/sites/test1/Shared Documents/abc", "TimeCreated": "2018-05-02T23:21:45Z", "TimeLastModified": "2018-05-02T23:21:45Z", "UniqueId": "0ac3da45-cacf-4c31-9b38-9ef3697d5a66", "WelcomePage": "" }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly pass params to request', (done) => {
    const request: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
        parentFolderUrl: '/Shared Documents',
        name: 'abc'
      }
    }, () => {
      try {
        assert(request.calledWith({ url: 'https://contoso.sharepoint.com/_api/web/folders',
        headers:
         { accept: 'application/json;odata=nometadata' },
        body: { ServerRelativeUrl: '/Shared Documents/abc' },
        json: true }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly pass params to request (sites/test1)', (done) => {
    const request: sinon.SinonStub = stubPostResponses();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com/sites/test1',
        parentFolderUrl: 'Shared Documents',
        name: 'abc'
      }
    }, () => {
      try {
        assert(request.calledWith({ url: 'https://contoso.sharepoint.com/sites/test1/_api/web/folders',
        headers:
         { accept: 'application/json;odata=nometadata' },
        body: { ServerRelativeUrl: '/sites/test1/Shared Documents/abc' },
        json: true }));
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
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', parentFolderUrl: '/Shared Documents', name: 'My Folder' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL and parentFolderUrl specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', parentFolderUrl: '/Shared Documents', name: 'My Folder' } });
    assert.strictEqual(actual, true);
  });
});