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
import command from './tenant-site-archive.js';

describe(commands.TENANT_SITE_ARCHIVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  const siteUrl = 'https://contoso.sharepoint.com/sites/Sales';
  const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                  <Actions>
                    <ObjectPath Id="2" ObjectPathId="1" />
                    <ObjectPath Id="4" ObjectPathId="3" />
                    <Query Id="5" ObjectPathId="3">
                      <Query SelectAllProperties="true">
                        <Properties />
                      </Query>
                    </Query>
                  </Actions>
                  <ObjectPaths>
                    <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                    <Method Id="3" ParentId="1" Name="ArchiveSiteByUrl">
                      <Parameters>
                        <Parameter Type="String">${siteUrl}</Parameter>
                      </Parameters>
                    </Method>
                  </ObjectPaths>
                  </Request>`;

  const requestResponse = JSON.stringify([
    {
      "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.24817.12008", "ErrorInfo": null, "TraceCorrelationId": "3bf525a1-202d-8000-b136-71cce6ed75ac"
    }, 2, {
      "IsNull": false
    },
    4,
    {
      "IsNull": false
    },
    5,
    {
      "_ObjectType_": "Microsoft.Online.SharePoint.TenantAdministration.SpoOperation",
      "_ObjectIdentity_": "3bf525a1-202d-8000-b136-71cce6ed75ac|908bed80-a04a-4433-b4a0-883d9847d110:1e852b49-bf4b-4ba5-bcd4-a8c4706c8ed4\nSpoOperation\nArchiveSite\n638505831432627104\nhttps%3a%2f%2fcontoso.sharepoint.com%2fsites%2fSales\n00000000-0000-0000-0000-000000000000",
      "HasTimedout": false,
      "IsComplete": true,
      "PollingInterval": 15000
    }
  ]);

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

    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TENANT_SITE_ARCHIVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid url passed', async () => {
    const actual = await command.validate({ options: { url: siteUrl } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before archiving the site when force option not passed', async () => {
    await command.action(logger, { options: { url: siteUrl } });

    assert(promptIssued);
  });

  it('aborts archiving the site when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();
    await command.action(logger, { options: { url: siteUrl } });
    assert(postStub.notCalled);
  });

  it('archives the site when prompt confirmed', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        return requestResponse;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);
    await command.action(logger, { options: { url: siteUrl } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, requestBody);
  });

  it('archives the site without prompting for confirmation when force option specified', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery`) {
        return requestResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: siteUrl, force: true, verbose: true } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, requestBody);
  });

  it('correctly handles API error', async () => {
    sinon.stub(request, 'post').resolves(JSON.stringify([
      {
        "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7303.1206", "ErrorInfo": {
          "ErrorMessage": "An error has occurred.", "ErrorValue": null, "TraceCorrelationId": "f70126a1-d0af-8000-c263-51051e36944e", "ErrorCode": -2147024894, "ErrorTypeName": "Microsoft.SharePoint.SPFieldValidationException"
        }, "TraceCorrelationId": "f70126a1-d0af-8000-c263-51051e36944e"
      }
    ]));

    await assert.rejects(command.action(logger, { options: { url: siteUrl, force: true } } as any),
      new CommandError('An error has occurred.'));
  });
});