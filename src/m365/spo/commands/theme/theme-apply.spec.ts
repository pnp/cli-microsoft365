import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import { spo } from '../../../../utils/spo';
import commands from '../../commands';
const command: Command = require('./theme-apply');

describe(commands.THEME_APPLY, () => {
  let log: string[];
  let requests: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    requests = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.THEME_APPLY);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('applies theme when correct parameters are passed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery', 'url');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC', 'request digest');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`, 'body');
    assert(loggerLogSpy.calledWith(true), 'log');
  });

  it('applies theme when correct parameters are passed (debug)', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
  });

  it('applies SharePoint theme (Blue) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Blue",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Orange) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Orange",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Red) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Red",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Purple) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Purple",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Green) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Green",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Gray) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Gray",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Dark Yellow) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Dark Yellow",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('applies SharePoint theme (Dark Blue) when correct parameters are passed', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);

      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "@odata.context": "https://contoso.sharepoint.com/sites/project-x/_api/$metadata#Edm.String",
          "value": "/sites/project-x/_catalogs/theme/Themed/6735E8EF"
        });
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "Dark Blue",
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    });
    let setRequestIssued = false;
    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        setRequestIssued = true;
      }
    });
    assert(setRequestIssued);
  });

  it('correctly handles error when applying custom theme', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    } as any), new CommandError('requestObjectIdentity ClientSvc error'));
  });

  it('handles unknown error command error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "ClientSvc unknown error" } }]);
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false
      }
    } as any), new CommandError('ClientSvc unknown error'));
  });

  it('handles command error correctly', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      requests.push(opts);
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc') {
          return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": "{ErrorMessage:error occurred}", "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, false]);
        }
      }

      if ((opts.url as string).indexOf(`/_api/ThemeManager/ApplyTheme`) > -1) {
        return JSON.stringify({
          "error": "Access denied. You do not have permission to perform this action or access this resource."
        });
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        name: 'Some color',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x',
        sharePointTheme: true
      }
    } as any), new CommandError('Access denied. You do not have permission to perform this action or access this resource.'));
    let correctRequestIssued = false;

    requests.forEach(r => {
      if (r.url.indexOf(`/_api/ThemeManager/ApplyTheme`) > -1 &&
        r.data) {
        correctRequestIssued = true;
      }
    });
    assert(correctRequestIssued);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        name: 'Some color',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    } as any), new CommandError('An error has occurred'));
  });

  it('passes validation when name is passed', async () => {
    const actual = await command.validate({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if webUrl is not passed', async () => {
    const actual = await command.validate({ options: { name: 'Contoso-Blue', webUrl: '' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { name: 'Contoso-Blue', webUrl: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when webUrl is passed', async () => {
    const actual = await command.validate({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if name is not a valid SharePoint theme name', async () => {
    const actual = await command.validate({ options: { name: 'invalid', webUrl: 'https://contoso.sharepoint.com/sites/project-x', sharePointTheme: true } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});
