import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./theme-apply');
import * as assert from 'assert';
import appInsights from '../../../../appInsights';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';
import config from '../../../../config';

describe(commands.THEME_APPLY, () => {
  let log: string[];
  let requests: any[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
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
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.THEME_APPLY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('applies theme when correct parameters are passed', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery', 'url');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC', 'request digest');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`, 'body');
        assert(cmdInstanceLogSpy.calledWith(true), 'log');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies theme when correct parameters are passed (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.strictEqual(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Blue) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Blue",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Orange) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Orange",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Red) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Red",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Purple) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Purple",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Green) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Green",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Gray) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Gray",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Dark Yellow) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Dark Yellow",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies SharePoint theme (Dark Blue) when correct parameters are passed', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: "Dark Blue",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let setRequestIssued = false;
      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          setRequestIssued = true;
        }
      });
      try {
        assert(setRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when applying custom theme', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown error command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "ClientSvc unknown error" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false,
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": "{ErrorMessage:error occurred}", "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]));
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return Promise.resolve(JSON.stringify({
          "error": {
            "code": "-2147024891, System.UnauthorizedAccessException",
            "message": "Access denied. You do not have permission to perform this action or access this resource."
          }
        }));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Some color',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    }, () => {
      let correctRequestIssued = false;

      requests.forEach(r => {
        if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
          r.body) {
          correctRequestIssued = true;
        }
      });
      try {
        assert(correctRequestIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => Promise.reject('An error has occurred'));

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Some color',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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

  it('passes validation when name is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if webUrl is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: '' } });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if name is not a valid SharePoint theme name', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'invalid', webUrl: 'https://contoso.sharepoint.com/sites/project-x', sharePointTheme: true } });
    assert.notStrictEqual(actual, true);
  });
});