import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command from './containertype-remove.js';
import { spe } from '../../../../utils/spe.js';
import { z } from 'zod';
import { CommandError } from '../../../../Command.js';
import config from '../../../../config.js';
import { spo } from '../../../../utils/spo.js';

describe(commands.CONTAINERTYPE_REMOVE, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;
  let confirmationPromptStub: sinon.SinonStub;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    auth.connection.active = true;
    auth.connection.spoUrl = spoAdminUrl.replace('-admin.sharepoint.com', '.sharepoint.com');
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso-admin.sharepoint.com' });
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
    confirmationPromptStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      spe.getContainerTypeIdByName,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both id and name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId, name: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither id nor name options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if id is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if id is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if name is passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerTypeName });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the container type', async () => {
    await command.action(logger, { options: { id: containerTypeId } });
    assert(confirmationPromptStub.calledOnce);
  });

  it('aborts removing the container type when prompt is not confirmed', async () => {
    const postStub = sinon.stub(request, 'post').resolves([]);

    await command.action(logger, { options: { name: containerTypeName } });
    assert(postStub.notCalled);
  });

  it('correctly removes a container type by id', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            SchemaVersion: '15.0.0.0',
            LibraryVersion: '16.0.25919.12007',
            ErrorInfo: null,
            TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
          },
          7,
          {
            IsNull: false
          }
        ];
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { id: containerTypeId, force: true, verbose: true } });
    assert.strictEqual(postStub.firstCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
    assert.strictEqual(postStub.lastCall.args[0].headers!['X-RequestDigest'], 'abc');
  });

  it('correctly removes a container type by name', async () => {
    sinon.stub(spe, 'getContainerTypeIdByName').resolves(containerTypeId);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            SchemaVersion: '15.0.0.0',
            LibraryVersion: '16.0.25919.12007',
            ErrorInfo: null,
            TraceCorrelationId: '864c91a1-f07a-c000-29c0-273ee30d83d8'
          },
          7,
          {
            IsNull: false
          }
        ];
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerTypeName, verbose: true, force: true } });
    assert.strictEqual(postStub.firstCall.args[0].data, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="7" ObjectPathId="6" /><Method Name="RemoveSPOContainerType" Id="8" ObjectPathId="6"><Parameters><Parameter TypeId="{b66ab1ca-fd51-44f9-8cfc-01f5c2a21f99}"><Property Name="ContainerTypeId" Type="Guid">{${containerTypeId}}</Property></Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="6" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
    assert.strictEqual(postStub.lastCall.args[0].headers!['X-RequestDigest'], 'abc');
  });

  it('correctly handles error when removing a container type', async () => {
    const errorMessage = `Tenant 7d858e1d-a366-48d1-8a15-ce45a733916b cannot delete Container Type ${containerTypeId} as it is a DirectToCustomer Container Type.`;

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `${spoAdminUrl}/_vti_bin/client.svc/ProcessQuery`) {
        return [
          {
            SchemaVersion: '15.0.0.0',
            LibraryVersion: '16.0.25919.12007',
            ErrorInfo: {
              ErrorMessage: errorMessage,
              ErrorValue: null,
              TraceCorrelationId: 'cd4a91a1-6041-c000-29c0-26f4566b5b74',
              ErrorCode: -2146232832,
              ErrorTypeName: 'Microsoft.SharePoint.SPException'
            },
            TraceCorrelationId: 'cd4a91a1-6041-c000-29c0-26f4566b5b74'
          }
        ];
      }

      throw 'Invalid request URL: ' + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { id: containerTypeId, force: true } }),
      new CommandError(errorMessage));
  });
});