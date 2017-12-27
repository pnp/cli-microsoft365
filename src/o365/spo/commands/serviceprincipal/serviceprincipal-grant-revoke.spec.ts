import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth, { Site } from '../../SpoAuth';
const command: Command = require('./serviceprincipal-grant-revoke');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import config from '../../../../config';
import Utils from '../../../../Utils';

describe(commands.SERVICEPRINCIPAL_GRANT_REVOKE, () => {
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
    assert.equal(command.name.startsWith(commands.SERVICEPRINCIPAL_GRANT_REVOKE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {}, url: 'https://contoso-admin.sharepoint.com' }, () => {
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
    cmdInstance.action({ options: {}, url: 'https://contoso-admin.sharepoint.com' }, () => {
      try {
        assert.equal(telemetry.name, commands.SERVICEPRINCIPAL_GRANT_REVOKE);
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

  it('aborts when not connected to a SharePoint tenant admin site', (done) => {
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError(`${auth.site.url} is not a tenant admin site. Connect to your tenant admin site and try again`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('revokes the specified permission grant (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, grantId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('revokes the specified permission grant', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers.authorization &&
        opts.headers.authorization.indexOf('Bearer ') === 0 &&
        opts.headers['X-RequestDigest'] &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectPath Id="12" ObjectPathId="11" /><ObjectPath Id="14" ObjectPathId="13" /><Method Name="DeleteObject" Id="15" ObjectPathId="13" /><Query Id="16" ObjectPathId="13"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="9" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /><Property Id="11" ParentId="9" Name="PermissionGrants" /><Method Id="13" ParentId="11" Name="GetByObjectId"><Parameters><Parameter Type="String">50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "63553a9e-101c-4000-d6f5-91ba841ffc9d"
          }, 66, {
            "IsNull": false
          }, 68, {
            "IsNull": false
          }, 70, {
            "IsNull": false
          }, 72, {
            "IsNull": false
          }, 73, {
            "IsNull": false
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, grantId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when revoking permission request', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": {
            "ErrorMessage": "The given key was not present in the dictionary.", "ErrorValue": null, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99", "ErrorCode": -1, "ErrorTypeName": "System.Collections.Generic.KeyNotFoundException"
          }, "TraceCorrelationId": "8da23a9e-00d0-4000-c621-0ffad6315d99"
        }
      ]));
    });
    auth.site = new Site();
    auth.site.connected = true;
    auth.site.url = 'https://contoso-admin.sharepoint.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(new CommandError('The given key was not present in the dictionary.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notEqual(typeof alias, 'undefined');
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

  it('allows specifying grantId', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--grantId') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the grantId option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation when the grantId is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { grantId: '50NAzUm3C0K9B6p8ORLtIvNe8tzf4ndKg51reFehHHg' } });
    assert(actual);
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
    assert(find.calledWith(commands.SERVICEPRINCIPAL_GRANT_REVOKE));
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
    auth.site.url = 'https://contoso-admin.sharepoint.com';
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