import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import config from '../../../../config';
import request from '../../../../request';
import { sinonUtil, spo } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./serviceprincipal-set');

describe(commands.SERVICEPRINCIPAL_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(spo, 'getRequestDigest').callsFake(() => Promise.resolve({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    }));
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      spo.getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SERVICEPRINCIPAL_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables the service principal (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, enabled: 'true', confirm: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          AccountEnabled: true,
          AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
          ReplyUrls: [
            "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables the service principal', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: false, enabled: 'true', confirm: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          AccountEnabled: true,
          AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
          ReplyUrls: [
            "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('disables the service principal (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">false</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": false, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
            ]
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { debug: true, enabled: 'false', confirm: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          AccountEnabled: false,
          AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
          ReplyUrls: [
            "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when approving permission request', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "InvalidOperationException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]));
    });
    command.action(logger, { options: { debug: false, confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before enabling service principal when confirmation argument not passed', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'true' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before disabling service principal when confirmation argument not passed', (done) => {
    command.action(logger, { options: { debug: false, enabled: 'false' } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts enabling service principal when prompt not confirmed', (done) => {
    const requestPostSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, enabled: 'true' } }, () => {
      try {
        assert(requestPostSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('enables the service principal when prompt confirmed', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.resolve(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
      }, 18, {
        "IsNull": false
      }, 21, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
          "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
        ]
      }
    ])));

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: false, enabled: 'true' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          AccountEnabled: true,
          AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
          ReplyUrls: [
            "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fa830edad9050849554e17113007.sharepoint.com\u002f"
          ]
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinon.stub(request, 'post').callsFake(() => Promise.reject('An error has occurred'));
    command.action(logger, { options: { debug: false, enabled: 'true', confirm: 'true' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('allows specifying the enabled option', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--enabled') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the enabled option is not a valid boolean value', async () => {
    const actual = await command.validate({ options: { enabled: '123' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the enabled option is true', async () => {
    const actual = await command.validate({ options: { enabled: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the enabled option is false', async () => {
    const actual = await command.validate({ options: { enabled: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});