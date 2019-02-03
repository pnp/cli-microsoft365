import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./sitedesign-run-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITEDESIGN_RUN_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.site = new Site();
    telemetry = null;
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
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SITEDESIGN_RUN_LIST), true);
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
        assert.equal(telemetry.name, commands.SITEDESIGN_RUN_LIST);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about site designs applied to the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
              "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
              "SiteDesignTitle": "Contoso Communication Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548959800000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
            "SiteDesignTitle": "Contoso Team Site",
            "StartTime": new Date(1548960114000).toLocaleString()
          },
          {
            "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
            "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
            "SiteDesignTitle": "Contoso Communication Site",
            "StartTime": new Date(1548959800000).toLocaleString()
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about the specified site design applied to the specified site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', siteDesignId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
            "SiteDesignTitle": "Contoso Team Site",
            "StartTime": new Date(1548960114000).toLocaleString()
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('outputs all information in JSON output mode', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
              "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
              "SiteDesignTitle": "Contoso Team Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548960114000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            },
            {
              "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
              "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
              "SiteDesignTitle": "Contoso Communication Site",
              "SiteDesignVersion": 1,
              "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
              "StartTime": "1548959800000",
              "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "b4411557-308b-4545-a3c4-55297d5cd8c8",
            "SiteDesignID": "6ec3ca5b-d04b-4381-b169-61378556d76e",
            "SiteDesignTitle": "Contoso Team Site",
            "SiteDesignVersion": 1,
            "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
            "StartTime": "1548960114000",
            "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
          },
          {
            "ID": "e15d5b37-fe95-4667-96f7-bee41aa1ccdf",
            "SiteDesignID": "2b5cb6bc-a176-472a-b59a-d1289d720414",
            "SiteDesignTitle": "Contoso Communication Site",
            "SiteDesignVersion": 1,
            "SiteID": "24cea241-ad89-44b8-8669-d60d88d38575",
            "StartTime": "1548959800000",
            "WebID": "e87e4ab8-2732-4a90-836d-9b3d0cd3a5cf"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving information about site designs', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('fails validation if webUrl not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if siteDesignId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', siteDesignId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if webUrl is valid and siteDesignId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.equal(actual, true);
  });

  it('passes validation if webUrl and siteDesignId are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', siteDesignId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } });
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
    assert(find.calledWith(commands.SITEDESIGN_RUN_LIST));
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
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a' } }, (err?: any) => {
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