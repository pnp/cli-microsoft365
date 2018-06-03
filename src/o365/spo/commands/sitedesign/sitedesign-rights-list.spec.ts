import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./sitedesign-rights-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';

describe(commands.SITEDESIGN_RIGHTS_LIST, () => {
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
    assert.equal(command.name.startsWith(commands.SITEDESIGN_RIGHTS_LIST), true);
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
        assert.equal(telemetry.name, commands.SITEDESIGN_RIGHTS_LIST);
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
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to a SharePoint Online site first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about permissions granted for the specified site design', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "1"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "View"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets information about permissions granted for the specified site design (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "1"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "View"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns original value for unknown permissions', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights`) > -1 &&
        JSON.stringify(opts.body) === JSON.stringify({
          id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b'
        })) {
        return Promise.resolve({
          "value": [
            {
              "DisplayName": "MOD Administrator",
              "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
              "Rights": "1"
            },
            {
              "DisplayName": "Patti Fernandez",
              "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
              "Rights": "2"
            }
          ]
        }
        );
      }

      return Promise.reject('Invalid request');
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "DisplayName": "MOD Administrator",
            "PrincipalName": "i:0#.f|membership|admin@contoso.onmicrosoft.com",
            "Rights": "View"
          },
          {
            "DisplayName": "Patti Fernandez",
            "PrincipalName": "i:0#.f|membership|pattif@contoso.onmicrosoft.com",
            "Rights": "2"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when site script not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: '0f27a016-d277-4bb4-b3c3-b5b040c9559b' } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if id not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } });
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
    assert(find.calledWith(commands.SITEDESIGN_RIGHTS_LIST));
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