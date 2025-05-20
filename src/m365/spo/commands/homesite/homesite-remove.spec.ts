import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
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
import command from './homesite-remove.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { z } from 'zod';

describe(commands.HOMESITE_REMOVE, () => {
  let log: any[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  const siteId = '00000000-0000-0000-0000-000000000010';
  const homeSites = {
    "value": [
      {
        "Audiences": [
          {
            "Email": "ColumnSearchable@contoso.onmicrosoft.com",
            "Id": "978b5280-4f80-47ea-a1db-b0d1d2fb1ba4",
            "Title": "ColumnSearchable Members"
          },
          {
            "Email": "contosoteam@contoso.onmicrosoft.com",
            "Id": "21af775d-17b3-4637-94a4-2ba8625277cb",
            "Title": "Contoso TeamR Members"
          }
        ],
        "IsInDraftMode": false,
        "IsVivaBackendSite": false,
        "SiteId": "431d7819-4aaf-49a1-b664-b2fe9e609b63",
        "TargetedLicenseType": 2,
        "Title": "The Landing",
        "Url": "https://contoso.sharepoint.com/sites/TheLanding",
        "VivaConnectionsDefaultStart": true,
        "WebId": "626c1724-8ac8-45d5-af87-c07c752fab75"
      },
      {
        "Audiences": [],
        "IsInDraftMode": false,
        "IsVivaBackendSite": false,
        "SiteId": "45d4a135-40e4-4571-8340-61d17fdfd58a",
        "TargetedLicenseType": 0,
        "Title": "Contoso Electronics",
        "Url": "https://contoso.sharepoint.com/sites/contosoportal",
        "VivaConnectionsDefaultStart": true,
        "WebId": "9418e2a1-855c-4752-8dd4-48693f43b10a"
      }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'ABC',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.promptForConfirmation,
      spo.getSiteAdminPropertiesByUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.HOMESITE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removing the Home Site when force option is not passed', async () => {
    await command.action(logger, { options: { debug: true } } as any);

    assert(promptIssued);
  });

  it('aborts removing Home Site when force option is not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'post');

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: {} });
    assert(postSpy.notCalled);
  });

  it('fails validation if the url is not a valid SharePoint url', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('removes the Home Site using legacy method when only one Home Site exists and prompt is confirmed', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return { value: [homeSites.value[0]] };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    let homeSiteRemoveCallIssued = false;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {

        homeSiteRemoveCallIssued = true;

        return JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been removed."
          ]
        );
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: {} });
    assert(homeSiteRemoveCallIssued);
  });

  it('removes the first Home Site when multiple Home Sites exist and url is not specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSites;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite` &&
        opts.data.siteId === siteId) {
        return {};
      }

      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been removed."
          ]
        );
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { force: true } });
    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].url, "https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite");
  });

  it('removes the Home Site specified by URL', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return homeSites;
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite` &&
        opts.data?.siteId === siteId) {
        return {};
      }

      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": null, "TraceCorrelationId": "e4f2e59e-c0a9-0000-3dd0-1d8ef12cc742"
            }, 57, {
              "IsNull": false
            }, 58, "The Home site has been removed."
          ]
        );
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com' } });
    assert(postStub.calledOnce);
  });

  it('correctly handles error when removing the Home Site (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/GetTargetedSitesDetails`) {
        return { value: [homeSites.value[0]] };
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.data === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="28" ObjectPathId="27" /><Method Name="RemoveSPHSite" Id="29" ObjectPathId="27" /></Actions><ObjectPaths><Constructor Id="27" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`) {
        return JSON.stringify(
          [
            {
              "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.8929.1227", "ErrorInfo": {
                "ErrorMessage": "The requested operation is part of an experimental feature that is not supported in the current environment.", "ErrorValue": null, "TraceCorrelationId": "75b6e89e-f072-8000-892f-75866252852a", "ErrorCode": -2146232832, "ErrorTypeName": "Microsoft.SharePoint.SPExperimentalFeatureException"
              }, "TraceCorrelationId": "f1f2e59e-3047-0000-3dd0-1f48be47bbc2"
            }
          ]
        );
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { debug: true, force: true } } as any),
      new CommandError(`The requested operation is part of an experimental feature that is not supported in the current environment.`));
  });
});
