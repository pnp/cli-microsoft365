import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spo } from '../../../../utils/spo.js';
import commands from '../../commands.js';
import command from './serviceprincipal-set.js';

describe(commands.SERVICEPRINCIPAL_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(Cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SERVICEPRINCIPAL_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('enables the service principal (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
            ]
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, enabled: true, force: true } });
    assert(loggerLogSpy.calledWith({
      AccountEnabled: true,
      AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
      ReplyUrls: [
        "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
      ]
    }));
  });

  it('enables the service principal', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">true</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
            ]
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { enabled: true, force: true } });
    assert(loggerLogSpy.calledWith({
      AccountEnabled: true,
      AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
      ReplyUrls: [
        "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
      ]
    }));
  });

  it('disables the service principal (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/_vti_bin/client.svc/ProcessQuery') > -1 &&
        opts.headers &&
        opts.headers['X-RequestDigest'] &&
        opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><SetProperty Id="29" ObjectPathId="27" Name="AccountEnabled"><Parameter Type="Boolean">false</Parameter></SetProperty><Method Name="Update" Id="30" ObjectPathId="27" /><Query Id="31" ObjectPathId="27"><Query SelectAllProperties="true"><Properties><Property Name="AccountEnabled" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Constructor Id="27" TypeId="{104e8f06-1e00-4675-99c6-1b9b504ed8d8}" /></ObjectPaths></Request>`) {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
          }, 18, {
            "IsNull": false
          }, 21, {
            "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": false, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
              "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
            ]
          }
        ]);
      }

      throw 'Invalid request';
    });
    await command.action(logger, { options: { debug: true, enabled: false, force: true } });
    assert(loggerLogSpy.calledWith({
      AccountEnabled: false,
      AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
      ReplyUrls: [
        "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
      ]
    }));
  });

  it('correctly handles error when approving permission request', async () => {
    sinon.stub(request, 'post').callsFake(async () => {
      return JSON.stringify([
        {
          "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
            "ErrorMessage": "An error has occurred", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "InvalidOperationException"
          }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
        }
      ]);
    });
    await assert.rejects(command.action(logger, { options: { force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('prompts before enabling service principal when confirmation argument not passed', async () => {
    await command.action(logger, { options: { enabled: true } });

    assert(promptIssued);
  });

  it('prompts before disabling service principal when confirmation argument not passed', async () => {
    await command.action(logger, { options: { enabled: false } });

    assert(promptIssued);
  });

  it('aborts enabling service principal when prompt not confirmed', async () => {
    const requestPostSpy = sinon.spy(request, 'post');
    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);
    await command.action(logger, { options: { enabled: true } });
    assert(requestPostSpy.notCalled);
  });

  it('enables the service principal when prompt confirmed', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7213.1200", "ErrorInfo": null, "TraceCorrelationId": "87b53a9e-90b1-4000-c0ac-27fb4ee21f41"
      }, 18, {
        "IsNull": false
      }, 21, {
        "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.Internal.SPOWebAppServicePrincipal", "AccountEnabled": true, "AppId": "57fb890c-0dab-4253-a5e0-7188c88b2bb4", "ReplyUrls": [
          "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
        ]
      }
    ]));

    sinonUtil.restore(Cli.promptForConfirmation);
    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { enabled: true } });
    assert(loggerLogSpy.calledWith({
      AccountEnabled: true,
      AppId: "57fb890c-0dab-4253-a5e0-7188c88b2bb4",
      ReplyUrls: [
        "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx?redirect", "https:\u002f\u002fcontoso.sharepoint.com\u002f_forms\u002fsinglesignon.aspx", "https:\u002f\u002fcontoso.sharepoint.com\u002f"
      ]
    }));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, { options: { enabled: true, force: true } } as any),
      new CommandError('An error has occurred'));
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
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

  it('passes validation when the enabled option is true', async () => {
    const actual = await command.validate({ options: { enabled: true } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the enabled option is false', async () => {
    const actual = await command.validate({ options: { enabled: false } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});