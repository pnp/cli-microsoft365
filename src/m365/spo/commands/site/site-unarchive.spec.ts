import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './site-unarchive.js';
import request from '../../../../request.js';
import config from '../../../../config.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';

describe(commands.SITE_UNARCHIVE, () => {
  const url = 'https://contoso.sharepoint.com/sites/project-x';
  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const response = '[{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.24817.12008","ErrorInfo":null,"TraceCorrelationId":"ab1127a1-5044-8000-b17c-4aafdd265386"},2,{"IsNull":false},4,{"IsNull":false}]';
  const requestBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">
                <Actions>
                  <ObjectPath Id="2" ObjectPathId="1" />
                  <ObjectPath Id="4" ObjectPathId="3" />
                </Actions>
                <ObjectPaths>
                  <Constructor Id="1" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" />
                  <Method Id="3" ParentId="1" Name="UnarchiveSiteByUrl">
                    <Parameters>
                      <Parameter Type="String">${url}</Parameter>
                    </Parameters>
                  </Method>
                </ObjectPaths>
              </Request>`;


  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    sinon.stub(spo, 'getRequestDigest').resolves({
      FormDigestValue: 'abc',
      FormDigestTimeoutSeconds: 1800,
      FormDigestExpiresAt: new Date(),
      WebFullUrl: 'https://contoso.sharepoint.com'
    });
    auth.connection.active = true;
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
      cli.promptForConfirmation,
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SITE_UNARCHIVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if url is not a valid SharePoint URL', async () => {
    const actual = await command.validate({ options: { url: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when a valid url is specified', async () => {
    const actual = await command.validate({ options: { url: url } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('aborts unarchiving site when prompt not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves();

    await command.action(logger, { options: { url: url } });
    assert(postStub.notCalled);
  });

  it('prompts before unarchiving the site when force option is not passed', async () => {
    await command.action(logger, { options: { url: url } });
    assert(promptIssued);
  });

  it('unarchives the site when prompt confirmed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: url, verbose: true } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, requestBody);
  });

  it('unarchives the site without prompting for confirmation when force option specified', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${adminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { url: url, force: true, verbose: true } });
    assert.deepStrictEqual(postStub.firstCall.args[0].data, requestBody);
  });

  it('correctly handles error when unarchiving site that does not exist', async () => {
    const errorMessage = 'File Not Found.';

    sinon.stub(request, 'post').resolves(`[{"SchemaVersion":"15.0.0.0","LibraryVersion":"16.0.24817.12008","ErrorInfo":{"ErrorMessage":"${errorMessage}","ErrorValue":null,"TraceCorrelationId":"731127a1-9041-8000-99fd-2865e0d78b49","ErrorCode":-2147024894,"ErrorTypeName":"System.IO.FileNotFoundException"},"TraceCorrelationId":"731127a1-9041-8000-99fd-2865e0d78b49"}]`);
    await assert.rejects(command.action(logger, { options: { url: url, force: true, verbose: true } }),
      new CommandError(errorMessage));
  });
});