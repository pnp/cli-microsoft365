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
import command, { options } from './storageentity-remove.js';

describe(commands.STORAGEENTITY_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

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
      spo.getTenantAppCatalogUrl
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.STORAGEENTITY_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes existing tenant property without prompting with confirmation argument', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ key: 'existingproperty', force: true, appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">existingproperty</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`);
  });

  it('removes tenant property using tenant app catalog URL when appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves('https://contoso.sharepoint.com/sites/appcatalog');

    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]);
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ key: 'existingproperty', force: true }) });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert(postStub.lastCall.args[0].data.includes('https://contoso.sharepoint.com/sites/appcatalog'));
  });

  it('throws error when tenant app catalog is not found and appCatalogUrl is not specified', async () => {
    sinon.stub(spo, 'getTenantAppCatalogUrl').resolves(null);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ key: 'existingproperty', force: true }) }),
      new CommandError('Tenant app catalog URL not found. Specify the URL of the app catalog site using the appCatalogUrl option.'));
  });

  it('prompts before removing tenant property when confirmation argument not passed', async () => {

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(promptIssued);
  });

  it('aborts removing property when prompt not confirmed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').rejects(new Error('Invalid request'));

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert(postStub.notCalled);
  });

  it('removes tenant property when prompt confirmed', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": null, "TraceCorrelationId": "4456299e-d09e-4000-ae61-ddde716daa27" }, 31, { "IsNull": false }, 33, { "IsNull": false }, 35, { "IsNull": false }]);
      }
      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'existingproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) });
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">existingproperty</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`);
  });

  it('correctly reports when trying to remove an nonexistent property', async () => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') {
        return JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7018.1204", "ErrorInfo": {
              "ErrorMessage": "File Not Found.", "ErrorValue": null, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26", "ErrorCode": -2147024894, "ErrorTypeName": "System.IO.FileNotFoundException"
            }, "TraceCorrelationId": "9e54299e-208a-4000-8546-cc4139091b26"
          }
        ]);
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'nonexistentproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) }), new CommandError('File Not Found.'));
    assert.strictEqual(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
    assert.strictEqual(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
    assert.strictEqual(postStub.lastCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="31" ObjectPathId="30" /><ObjectPath Id="33" ObjectPathId="32" /><ObjectPath Id="35" ObjectPathId="34" /><Method Name="RemoveStorageEntity" Id="36" ObjectPathId="34"><Parameters><Parameter Type="String">nonexistentproperty</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="30" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="32" ParentId="30" Name="GetSiteByUrl"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/appcatalog</Parameter></Parameters></Method><Property Id="34" ParentId="32" Name="RootWeb" /></ObjectPaths></Request>`);
  });

  it('fails validation if appCatalogUrl is not a valid URL', () => {
    const actual = commandOptionsSchema.safeParse({ appCatalogUrl: 'foo', key: 'prop' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation when appCatalogUrl is a valid SharePoint URL', () => {
    const actual = commandOptionsSchema.safeParse({ appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog', key: 'prop' });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation when appCatalogUrl is not specified', () => {
    const actual = commandOptionsSchema.safeParse({ key: 'prop' });
    assert.strictEqual(actual.success, true);
  });

  it('handles promise rejection', async () => {
    sinonUtil.restore(spo.getRequestDigest);
    sinon.stub(spo, 'getRequestDigest').rejects(new Error('getRequestDigest error'));

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ debug: true, key: 'nonexistentproperty', appCatalogUrl: 'https://contoso.sharepoint.com/sites/appcatalog' }) }), new CommandError('getRequestDigest error'));
  });
});
