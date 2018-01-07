import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./sitedesign-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITEDESIGN_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
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
      auth.ensureAccessToken,
      auth.restoreAuth,
      (command as any).getRequestDigest
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SITEDESIGN_ADD), true);
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
        assert.equal(telemetry.name, commands.SITEDESIGN_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to a SharePoint site', (done) => {
    auth.site = new Site();
    auth.site.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site design for a team site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site design for a team site (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design wilt multiple site script IDs', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24', '449c0c6d-5380-4df2-b84b-622e0ac8ec25']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24", "449c0c6d-5380-4df2-b84b-622e0ac8ec25"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24, 449c0c6d-5380-4df2-b84b-622e0ac8ec25" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledOnce);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new site design for a communication site', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '68',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24']
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 68
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'CommunicationSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 68
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with description', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site'
          }
        })) {
        return Promise.resolve({
          "Description": "Contoso team site",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": "Contoso team site",
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with previewImageUrl', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageUrl: 'https://contoso.com/assets/team-site-preview.png' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": null,
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with previewImageAltText', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            PreviewImageAltText: 'Contoso team site preview'
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", previewImageAltText: 'Contoso team site preview' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": false,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new default team site site design', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            IsDefault: true
          }
        })) {
        return Promise.resolve({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", isDefault: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": null,
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": null,
          "PreviewImageUrl": null,
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds new team site site design with all options specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          info: {
            Title: 'Contoso',
            WebTemplate: '64',
            SiteScriptIds: ['449c0c6d-5380-4df2-b84b-622e0ac8ec24'],
            Description: 'Contoso team site',
            PreviewImageUrl: 'https://contoso.com/assets/team-site-preview.png',
            PreviewImageAltText: 'Contoso team site preview',
            IsDefault: true
          }
        })) {
        return Promise.resolve({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24", description: 'Contoso team site', previewImageUrl: 'https://contoso.com/assets/team-site-preview.png', previewImageAltText: 'Contoso team site preview', isDefault: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "Description": 'Contoso team site',
          "Id": "2a9f178a-4d1d-449c-9296-df509ab4702c",
          "IsDefault": true,
          "PreviewImageAltText": 'Contoso team site preview',
          "PreviewImageUrl": 'https://contoso.com/assets/team-site-preview.png',
          "SiteScriptIds": ["449c0c6d-5380-4df2-b84b-622e0ac8ec24"],
          "Title": "Contoso",
          "Version": 1,
          "WebTemplate": 64
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles OData error when creating site script', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'An error has occurred' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, title: 'Contoso', webTemplate: 'TeamSite', siteScripts: '449c0c6d-5380-4df2-b84b-622e0ac8ec24' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('An error has occurred')));
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

  it('supports specifying title', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--title') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying webTemplate', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--webTemplate') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying siteScripts', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--siteScripts') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying description', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--description') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageUrl', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageUrl') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying previewImageAltText', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--previewImageAltText') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying isDefault', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--isDefault') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if title not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if webTemplate not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified webTemplate is invalid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'Invalid' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if siteScripts not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'TeamSite' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if specified siteScripts is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if the second specified siteScriptId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,abc" } });
    assert.notEqual(actual, true);
  });

  it('passes validation if all required parameters are valid', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24" } });
    assert.equal(actual, true);
  });

  it('passes validation if all required parameters are valid (multiple siteScripts)', () => {
    const actual = (command.validate() as CommandValidate)({ options: { title: 'Contoso', webTemplate: 'TeamSite', siteScripts: "449c0c6d-5380-4df2-b84b-622e0ac8ec24,449c0c6d-5380-4df2-b84b-622e0ac8ec25" } });
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
    assert(find.calledWith(commands.SITEDESIGN_ADD));
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
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});