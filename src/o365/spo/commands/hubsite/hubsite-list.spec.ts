import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./hubsite-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.HUBSITE_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
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
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.HUBSITE_LIST), true);
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
        assert.equal(telemetry.name, commands.HUBSITE_LIST);
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

  it('lists hub sites', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
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
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Title": "Sales"
          },
          {
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists hub sites (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
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
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Title": "Sales"
          },
          {
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists hub sites with all properties for JSON output', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/hubsites`) > -1) {
        return Promise.resolve({
          value: [
            {
              "Description": null,
              "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
              "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Sales"
            },
            {
              "Description": null,
              "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
              "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
              "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
              "Targets": null,
              "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
              "Title": "Travel Programs"
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
    cmdInstance.action({ options: { debug: false, output: 'json' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "Description": null,
            "ID": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "389d0d83-40bb-40ad-b92a-534b7cb37d0b",
            "SiteUrl": "https://contoso.sharepoint.com/sites/Sales",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Sales"
          },
          {
            "Description": null,
            "ID": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "LogoUrl": "http://contoso.com/__siteIcon__.jpg",
            "SiteId": "b2c94ca1-0957-4bdd-b549-b7d365edc10f",
            "SiteUrl": "https://contoso.sharepoint.com/sites/travelprograms",
            "Targets": null,
            "TenantInstanceId": "00000000-0000-0000-0000-000000000000",
            "Title": "Travel Programs"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when retrieving available site designs', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
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
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
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
    assert(find.calledWith(commands.HUBSITE_LIST));
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
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
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