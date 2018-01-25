import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./web-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.WEB_ADD, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;
 
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
      request.get,
      request.post
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
    assert.equal(command.name.startsWith(commands.WEB_ADD), true);
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
        assert.equal(telemetry.name, commands.WEB_ADD);
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
    cmdInstance.action({ options: { debug: false, webUrl: 'https://contoso.sharepoint.com', title: 'Documents' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Connect to a SharePoint Online site first')));
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.WEB_ADD));
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
        title: "subsite",
        webUrl: "subsite",
        parentWebUrl:"https://contoso.sharepoint.com",
        debug: false
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the title option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  webUrl: "subsite",
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('passes validation if all the options are specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", webUrl: "subsite",
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.equal(actual, true);
  });

  it('fails validation if the weburl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the parentWeburl option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    webUrl:"subsite", webTemplate:"STS#0", locale:1033} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the webTemplate option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    webUrl:"subsite", parentWebUrl:"https://contoso.sharepoint.com", locale:1033} });
    assert.notEqual(actual, true);
  });


  it('fails validation if the locale option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {  title:"subsite", 
    webUrl:"subsite", parentWebUrl:"https://contoso.sharepoint.com", webTemplate:"STS#0"} });
    assert.notEqual(actual, true);
  });

  it('creates web and inherits the navigation', (done) => {
    Utils.restore(auth.getAccessToken);
    Utils.restore(sinon.stub);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }

      return Promise.resolve('abc');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/effectivebasepermissions') > -1) {
        return Promise.resolve(
          { High:2147483647,
            Low:4294967295
          }
        );
      }

      return Promise.resolve('abc');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : true,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(`Subsite subsite created.`));
      assert(cmdInstanceLogSpy.calledWith("Setting the navigation to inherit the parent settings."));
      done();
    });
  });

  it('creates web only without inheriting the navigation', (done) => {
    Utils.restore(auth.getAccessToken);
    Utils.restore(sinon.stub);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('_api/web/webinfos/add') > -1) {
        return Promise.resolve({ Configuration: 0,
          Created                 : "2018-01-24T18:24:20",
          Description             : "subsite",
          Id                      : "08385b9a-8d5f-4ee9-ac98-bf6984c1856b",
          Language                : 1033,
          LastItemModifiedDate    : "2018-01-24T18:24:27Z",
          LastItemUserModifiedDate: "2018-01-24T18:24:27Z",
          ServerRelativeUrl       : "/subsite",
          Title                   : "subsite",
          WebTemplate             : "STS",
          WebTemplateId           : 0 });
      }

      return Promise.resolve('abc');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({  options: {
      title: "subsite",
      webUrl: "subsite",
      parentWebUrl:"https://contoso.sharepoint.com",
      inheritNavigation : false,
      local:1033,
      debug: true
    } }, () => {
      assert(cmdInstanceLogSpy.calledWith(`Subsite subsite created.`));
      assert.equal(cmdInstanceLogSpy.calledWith("Setting the navigation to inherit the parent settings."), false, "Should not set the inheritnavigation.");
      done();
    });
  });

});