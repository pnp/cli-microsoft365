import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./sitedesign-run-status-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITEDESIGN_RUN_STATUS_GET, () => {
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
    assert.equal(command.name.startsWith(commands.SITEDESIGN_RUN_STATUS_GET), true);
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
        assert.equal(telemetry.name, commands.SITEDESIGN_RUN_STATUS_GET);
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
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
        return Promise.resolve({
          "value": [
            { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
            { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ActionTitle": "Add to hub site",
            "SiteScriptTitle": "Contoso Team Site",
            "OutcomeText": "One or more of the properties on this action has an invalid type."
          },
          {
            "ActionTitle": "Associate SPFX extension Collab Footer",
            "SiteScriptTitle": "Contoso Team Site",
            "OutcomeText": null
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
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus`) > -1) {
        return Promise.resolve({
          "value": [
            { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
            { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8', output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          { "ActionIndex": 0, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Add to hub site", "LastModified": "1548960114000", "OrdinalIndex": 0, "OutcomeCode": 1, "OutcomeText": "One or more of the properties on this action has an invalid type.", "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" },
          { "ActionIndex": 1, "ActionKey": "00000000-0000-0000-0000-000000000000", "ActionTitle": "Associate SPFX extension Collab Footer", "LastModified": "1548960114000", "OrdinalIndex": 1, "OutcomeCode": 0, "OutcomeText": null, "SiteScriptID": "f37c6396-97fa-4fff-9d7e-3ed44faaf608", "SiteScriptIndex": 0, "SiteScriptTitle": "Contoso Team Site" }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when the specified runId doesn\'t point to a valid run', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'Value does not fall within the expected range' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Value does not fall within the expected range')));
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
    const actual = (command.validate() as CommandValidate)({ options: { runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'invalid', runId: 'b4411557-308b-4545-a3c4-55297d5cd8c8' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if runId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if runId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', runId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if webUrl and runId are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { webUrl: 'https://contoso.sharepoint.com', runId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } });
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
    assert(find.calledWith(commands.SITEDESIGN_RUN_STATUS_GET));
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com/sites/team-a', runId: '6ec3ca5b-d04b-4381-b169-61378556d76e' } }, (err?: any) => {
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