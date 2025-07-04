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

  it('fails validation if the url option is not a valid SharePoint site url', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the url option is a valid SharePoint site URL', async () => {
    const actual = commandOptionsSchema.safeParse({ url: 'https://contoso.sharepoint.com' });
    assert.strictEqual(actual.success, true);
  });

  it('removes the Home Site when prompt confirmed', async () => {
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

  it('removes the Home Site whithout confirm prompt', async () => {
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

    await command.action(logger, { options: { force: true } });
    assert(homeSiteRemoveCallIssued);
  });

  it('removes the Home Site specified by URL', async () => {
    sinon.stub(spo, 'getSiteAdminPropertiesByUrl').resolves({ SiteId: siteId } as any);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_api/SPO.Tenant/RemoveTargetedSite`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { url: 'https://contoso.sharepoint.com', verbose: true } });
    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.lastCall.args[0].data, { siteId });
  });

  it('correctly handles error when removing the Home Site (debug)', async () => {
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

  it('correctly handles error when attempting to remove a site that is not a home site or Viva Connections', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        "odata.error": {
          "code": "-2146232832, Microsoft.SharePoint.SPException",
          "message": {
            "lang": "en-US",
            "value": "[Error ID: 03fc404e-0f70-4607-82e8-8fdb014e8658] The site with ID \"8e4686ed-b00c-4c5f-a0e2-4197081df5d5\" has not been added as a home site or Viva Connections. Check aka.ms/homesites for details."
          }
        }
      }
    });

    await assert.rejects(
      command.action(logger, { options: { debug: true, force: true } } as any),
      new CommandError('[Error ID: 03fc404e-0f70-4607-82e8-8fdb014e8658] The site with ID \"8e4686ed-b00c-4c5f-a0e2-4197081df5d5\" has not been added as a home site or Viva Connections. Check aka.ms/homesites for details.')
    );
  });
});
