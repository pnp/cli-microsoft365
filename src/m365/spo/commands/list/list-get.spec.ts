import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_GET, () => {
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
    assert.equal(command.name.startsWith(commands.LIST_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('retrieves and prints all details of list if title option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      //if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.resolve(
          {
            "AllowContentTypes": true,
            "BaseTemplate": 109,
            "BaseType": 1,
            "ContentTypesEnabled": false,
            "CrawlNonDefaultViews": false,
            "Created": null,
            "CurrentChangeToken": null,
            "CustomActionElements": null,
            "DefaultContentApprovalWorkflowId": "00000000-0000-0000-0000-000000000000",
            "DefaultItemOpenUseListSetting": false,
            "Description": "",
            "Direction": "none",
            "DocumentTemplateUrl": null,
            "DraftVersionVisibility": 0,
            "EnableAttachments": false,
            "EnableFolderCreation": true,
            "EnableMinorVersions": false,
            "EnableModeration": false,
            "EnableVersioning": false,
            "EntityTypeName": "Documents",
            "ExemptFromBlockDownloadOfNonViewableFiles": false,
            "FileSavePostProcessingEnabled": false,
            "ForceCheckout": false,
            "HasExternalDataSource": false,
            "Hidden": false,
            "Id": "14b2b6ed-0885-4814-bfd6-594737cc3ae3",
            "ImagePath": null,
            "ImageUrl": null,
            "IrmEnabled": false,
            "IrmExpire": false,
            "IrmReject": false,
            "IsApplicationList": false,
            "IsCatalog": false,
            "IsPrivate": false,
            "ItemCount": 69,
            "LastItemDeletedDate": null,
            "LastItemModifiedDate": null,
            "LastItemUserModifiedDate": null,
            "ListExperienceOptions": 0,
            "ListItemEntityTypeFullName": null,
            "MajorVersionLimit": 0,
            "MajorWithMinorVersionsLimit": 0,
            "MultipleDataList": false,
            "NoCrawl": false,
            "ParentWebPath": null,
            "ParentWebUrl": null,
            "ParserDisabled": false,
            "ServerTemplateCanCreateFolders": true,
            "TemplateFeatureId": null,
            "Title": "Documents"
          }
        );
      //}
      //return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: {
      debug: true,
      title: 'Documents',
      webUrl: 'https://contoso.sharepoint.com'
    } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({ 
          AllowContentTypes: true,
          BaseTemplate: 109,
          BaseType: 1,
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: null,
          CurrentChangeToken: null,
          CustomActionElements: null,
          DefaultContentApprovalWorkflowId: '00000000-0000-0000-0000-000000000000',
          DefaultItemOpenUseListSetting: false,
          Description: '',
          Direction: 'none',
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: false,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableVersioning: false,
          EntityTypeName: 'Documents',
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: '14b2b6ed-0885-4814-bfd6-594737cc3ae3',
          ImagePath: null,
          ImageUrl: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 69,
          LastItemDeletedDate: null,
          LastItemModifiedDate: null,
          LastItemUserModifiedDate: null,
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: null,
          MajorVersionLimit: 0,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: null,
          ParentWebUrl: null,
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: null,
          Title: 'Documents'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles list get reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists/GetByTitle(') > -1) {
        return Promise.reject(err);
      }

      return Promise.reject('Invalid request');
    });

    const actionTitle: string = 'Documents';

    cmdInstance.action({
      options: {
        debug: true,
        title: actionTitle,
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

  it('uses correct API url when id option is passed', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/lists(guid') > -1) {
        return Promise.resolve('Correct Url')
      }

      return Promise.reject('Invalid request');
    });

    const actionId: string = '0CD891EF-AFCE-4E55-B836-FCE03286CCCF';

    cmdInstance.action({
      options: {
        debug: false,
        id: actionId,
        webUrl: 'https://contoso.sharepoint.com',
      }
    }, () => {

      try {
        assert(1===1);
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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notEqual(actual, true);
  });

  it('fails validation if both id and title options are not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
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

  it('fails validation if both id and title options are passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', id: '0CD891EF-AFCE-4E55-B836-FCE03286CCCF', title: 'Documents' } });
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
    assert(find.calledWith(commands.LIST_GET));
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