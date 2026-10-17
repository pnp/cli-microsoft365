import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import command, { options } from './site-appcatalog-remove.js';

describe(commands.SITE_APPCATALOG_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');

    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });
    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.getSettingWithDefaultValue,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_APPCATALOG_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the site app catalog when force option not passed', async () => {
    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/site' }) });
    assert(promptIssued);
  });

  it('aborts removing site app catalog when prompt not confirmed', async () => {
    const postSpy = sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/site' }) });
    assert(postSpy.notCalled);
  });

  it('removes site collection app catalog (debug=false) without prompting for confirmation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
            }, 50, {
              "IsNull": false
            }, 52, {
              "IsNull": false
            }, 54, {
              "IsNull": false
            }, 56, {
              "IsNull": false
            }, 58, {
              "IsNull": false
            }
          ]
          );
        }
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ siteUrl: 'https://contoso.sharepoint.com/sites/site', force: true }) });
  });

  it('removes site collection app catalog (debug=true) while prompting for confirmation', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": null, "TraceCorrelationId": "7cd0389e-6015-4000-979e-22c0a7af5f43"
            }, 50, {
              "IsNull": false
            }, 52, {
              "IsNull": false
            }, 54, {
              "IsNull": false
            }, 56, {
              "IsNull": false
            }, 58, {
              "IsNull": false
            }
          ]
          );
        }
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/site' }) });
    assert(loggerLogToStderrSpy.called);
  });

  it('correctly handles error when removing site collection app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        if (opts.headers &&
          opts.headers['X-RequestDigest'] &&
          opts.headers['X-RequestDigest'] === 'abc' &&
          opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="50" ObjectPathId="49" /><ObjectPath Id="52" ObjectPathId="51" /><ObjectPath Id="54" ObjectPathId="53" /><ObjectPath Id="56" ObjectPathId="55" /><ObjectPath Id="58" ObjectPathId="57" /><Method Name="Remove" Id="59" ObjectPathId="57"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="49" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="51" ParentId="49" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/site</Parameter></Parameters></Method><Property Id="53" ParentId="51" Name="RootWeb" /><Property Id="55" ParentId="53" Name="TenantAppCatalog" /><Property Id="57" ParentId="55" Name="SiteCollectionAppCatalogsSites" /></ObjectPaths></Request>`) {
          return JSON.stringify([
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7206.1203", "ErrorInfo": {
                "ErrorMessage": "Unknown Error", "ErrorValue": null, "TraceCorrelationId": "d9d0389e-002f-4000-b24b-dedcfcca7277", "ErrorCode": -1, "ErrorTypeName": "Microsoft.SharePoint.Client.UnknownError"
              }, "TraceCorrelationId": "d9d0389e-002f-4000-b24b-dedcfcca7277"
            }
          ]
          );
        }
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/site', force: true }) }), new CommandError('Unknown Error'));
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, siteUrl: 'https://contoso.sharepoint.com/sites/site', force: true }) }), new CommandError('An error has occurred'));
  });

  it('passes validation when the url option specified', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/site'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation when the url option not specified', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when the url option is not a valid SharePoint site URL', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'foo'
    });
    assert.strictEqual(actual.success, false);
    assert.strictEqual(actual.error?.issues[0].message, `'foo' is not a valid SharePoint Online site URL.`);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      siteUrl: 'https://contoso.sharepoint.com/sites/site',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });
});
