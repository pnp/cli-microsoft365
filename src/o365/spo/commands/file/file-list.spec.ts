import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./file-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.FILE_LIST, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
  let stubAuth: any = () => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/common/oauth2/token') > -1) {
        return Promise.resolve('abc');
      }

      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });
  }

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'getAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => { return { FormDigestValue: 'abc' }; });
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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.getAccessToken,
      auth.restoreAuth,
      request.get
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.FILE_LIST), true);
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
        assert.equal(telemetry.name, commands.FILE_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all files', (done) => {
    stubAuth();
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {"value":[{
            "CheckInComment": "",
            "CheckOutType": 2,
            "ContentTag": "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
            "CustomizedPageStatus": 0,
            "ETag": "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
            "Exists": true,
            "IrmEnabled": false,
            "Length": "331673",
            "Level": 1,
            "LinkingUri": "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
            "LinkingUrl": "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
            "MajorVersion": 3,
            "MinorVersion": 0,
            "Name": "Test.docx",
            "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx",
            "TimeCreated": "2018-02-05T08:42:36Z",
            "TimeLastModified": "2018-02-05T08:44:03Z",
            "Title": "",
            "UIVersion": 1536,
            "UIVersionLabel": "3.0",
            "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b"
          }]}
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {
      output: 'json',
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      folder: 'Shared Documents'
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({value: [{ 
          CheckInComment: "",
          CheckOutType: 2,
          ContentTag: "{F09C4EFE-B8C0-4E89-A166-03418661B89B},9,12",
          CustomizedPageStatus: 0,
          ETag: "\"{F09C4EFE-B8C0-4E89-A166-03418661B89B},9\"",
          Exists: true,
          IrmEnabled: false,
          Length: "331673",
          Level: 1,
          LinkingUri: "https://contoso.sharepoint.com/sites/project-x/Shared%20documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
          LinkingUrl: "https://contoso.sharepoint.com/sites/project-x/Shared Documents/Test.docx?d=wf09c4efeb8c04e89a16603418661b89b",
          MajorVersion: 3,
          MinorVersion: 0,
          Name: "Test.docx",
          ServerRelativeUrl: "/sites/project-x/Shared documents/Test.docx",
          TimeCreated: "2018-02-05T08:42:36Z",
          TimeLastModified: "2018-02-05T08:44:03Z",
          Title: "",
          UIVersion: 1536,
          UIVersionLabel: "3.0",
          UniqueId: "f09c4efe-b8c0-4e89-a166-03418661b89b"
        }]}));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('retrieves all files with output option text', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.resolve(
          {"value":[
            {
              "UniqueId": "f09c4efe-b8c0-4e89-a166-03418661b89b",
              "Name": "Test.docx",
              "ServerRelativeUrl": "/sites/project-x/Shared documents/Test.docx"
            }
          ]}
        );
      }
      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    auth.site.tenantId = 'abc';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {
      output: 'text',
      debug: false,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      folder: 'Shared Documents'
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [{
            UniqueId: 'f09c4efe-b8c0-4e89-a166-03418661b89b',
            Name: 'Test.docx',
            ServerRelativeUrl: '/sites/project-x/Shared documents/Test.docx'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore(request.get);
        Utils.restore(request.post);
      }
    });
  });

  it('command correctly handles files list reject request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_api/contextinfo') > -1) {
        return Promise.resolve({
          FormDigestValue: 'abc'
        });
      }

      return Promise.reject('Invalid request');
    });

    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        debug: true,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, (error?: any) => {
      try {
        assert.equal(JSON.stringify(error), JSON.stringify(new CommandError(err)));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
        ]);
      }
    });
  });

  it('uses correct API url when output json option is passed', (done) => {
    stubAuth();

    sinon.stub(request, 'get').callsFake((opts) => {
      cmdInstance.log('Test Url:');
      cmdInstance.log(opts.url);
      if (opts.url.indexOf('select123=') > -1) {
        return Promise.resolve('Correct Url1')
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();

    cmdInstance.action({
      options: {
        output: 'json',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com',
        folder: 'Shared Documents'
      }
    }, () => {

      try {
        assert('Correct Url');
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        Utils.restore([
          request.post,
          request.get
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

  it('fails validation if the webUrl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { folder: '/' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo', folder: '/' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the webUrl option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', folder: '/' } });
    assert.equal(actual, true);
  });

  it('fails validation if the folder option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
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
    assert(find.calledWith(commands.FILE_LIST));
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
        webUrl: "https://contoso.sharepoint.com",
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});