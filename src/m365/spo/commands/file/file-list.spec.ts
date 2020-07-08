import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./file-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.FILE_LIST, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.FILE_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('retrieves all files', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
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

    cmdInstance.action({ options: {
      output: 'json',
      debug: true,
      webUrl: 'https://contoso.sharepoint.com/sites/project-x',
      folder: 'Shared Documents'
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([{ 
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
        }]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all files with output option text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
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
    });
  });

  it('command correctly handles files list reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/GetFolderByServerRelativeUrl') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

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
    });
  });

  it('uses correct API url when output json option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('select123=') > -1) {
        return Promise.resolve('Correct Url1')
      }

      return Promise.reject('Invalid request');
    });

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
});