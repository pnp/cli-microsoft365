import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./web-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';

describe(commands.WEB_LIST, () => {
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
    assert.equal(command.name.startsWith(commands.WEB_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('retrieves all webs', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
        return Promise.resolve(
          {
            "value": [{
              "AllowRssFeeds": false,
              "AlternateCssUrl": null,
              "AppInstanceId": "00000000-0000-0000-0000-000000000000",
              "Configuration": 0,
              "Created": null,
              "CurrentChangeToken": null,
              "CustomMasterUrl": null,
              "Description": null,
              "DesignPackageId": null,
              "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled": false,
              "EnableMinimalDownload": false,
              "HorizontalQuickLaunch": false,
              "Id": "d8d179c7-f459-4f90-b592-14b08e84accb",
              "IsMultilingual": false,
              "Language": 1033,
              "LastItemModifiedDate": null,
              "LastItemUserModifiedDate": null,
              "MasterUrl": null,
              "NoCrawl": false,
              "OverwriteTranslationsOnChange": false,
              "ResourcePath": null,
              "QuickLaunchEnabled": false,
              "RecycleBinEnabled": false,
              "ServerRelativeUrl": null,
              "SiteLogoUrl": null,
              "SyndicationEnabled": false,
              "Title": "Subsite",
              "TreeViewEnabled": false,
              "UIVersion": 15,
              "UIVersionConfigurationEnabled": false,
              "Url": "https://Contoso.sharepoint.com/Subsite",
              "WebTemplate": "STS",
            }]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'json',
        debug: true,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          value: [{
            AllowRssFeeds: false,
            AlternateCssUrl: null,
            AppInstanceId: "00000000-0000-0000-0000-000000000000",
            Configuration: 0,
            Created: null,
            CurrentChangeToken: null,
            CustomMasterUrl: null,
            Description: null,
            DesignPackageId: null,
            DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
            EnableMinimalDownload: false,
            HorizontalQuickLaunch: false,
            Id: "d8d179c7-f459-4f90-b592-14b08e84accb",
            IsMultilingual: false,
            Language: 1033,
            LastItemModifiedDate: null,
            LastItemUserModifiedDate: null,
            MasterUrl: null,
            NoCrawl: false,
            OverwriteTranslationsOnChange: false,
            ResourcePath: null,
            QuickLaunchEnabled: false,
            RecycleBinEnabled: false,
            ServerRelativeUrl: null,
            SiteLogoUrl: null,
            SyndicationEnabled: false,
            Title: "Subsite",
            TreeViewEnabled: false,
            UIVersion: 15,
            UIVersionConfigurationEnabled: false,
            Url: "https://Contoso.sharepoint.com/Subsite",
            WebTemplate: "STS",
          }]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves all webs with output option text', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
        return Promise.resolve(
          {
            "value": [
              {
                "Title": "Subsite",
                "Url": "https://Contoso.sharepoint.com/",
                "Id": "d8d179c7-f459-4f90-b592-14b08e84accb"
              }
            ]
          }
        );
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        output: 'text',
        debug: false,
        webUrl: 'https://contoso.sharepoint.com'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [{
            Title: 'Subsite',
            Url: "https://Contoso.sharepoint.com/",
            Id: 'd8d179c7-f459-4f90-b592-14b08e84accb'
          }]
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('command correctly handles web list reject request', (done) => {
    const err = 'Invalid request';
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_api/web/webs') > -1) {
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
      cmdInstance.log('Test Url:');
      cmdInstance.log(opts.url);
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

  it('fails validation if the url option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the url option is not a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'foo' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the url option is a valid SharePoint site URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.WEB_LIST));
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