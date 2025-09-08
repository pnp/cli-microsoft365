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
import command from './containertype-get.js';
import { CommandError } from '../../../../Command.js';
import { z } from 'zod';
import { formatting } from '../../../../utils/formatting.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.CONTAINERTYPE_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: z.ZodTypeAny;

  const containerTypeId = '3ec7c59d-ef31-0752-1ab5-5c343a5e8557';
  const containerTypeName = 'SharePoint Embedded Free Trial Container Type';
  const containerTypeResponse = {
    "id": containerTypeId,
    "name": containerTypeName,
    "owningAppId": "11335700-9a00-4c00-84dd-0c210f203f00",
    "billingClassification": "trial",
    "billingStatus": "valid",
    "createdDateTime": "01/20/2025",
    "expirationDateTime": "02/20/2025",
    "etag": "RVRhZw==",
    "settings": {
      "urlTemplate": "https://app.contoso.com/redirect?tenant={tenant-id}&drive={drive-id}&folder={folder-id}&item={item-id}",
      "isDiscoverabilityEnabled": true,
      "isSearchEnabled": true,
      "isItemVersioningEnabled": true,
      "itemMajorVersionLimit": 50,
      "maxStoragePerContainerInBytes": 104857600,
      "isSharingRestricted": false,
      "consumingTenantOverridables": ""
    }
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.spoUrl = 'https://contoso.sharepoint.com';
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse()!;
    sinon.stub(accessToken, 'assertAccessTokenType').withArgs('delegated').returns();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINERTYPE_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').rejects({
      error: {
        code: 'itemNotFound',
        message: 'Item not found',
        innerError: {
          date: '2025-08-27T11:45:11',
          'request-id': '53ac080f-14a3-4bd0-be8a-b78150c1e11d',
          'client-request-id': '53ac080f-14a3-4bd0-be8a-b78150c1e11d'
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { id: containerTypeId, verbose: true } }),
      new CommandError('Item not found'));
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: '123' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if the id is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ id: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('passes validation if required options specified (displayName)', async () => {
    const actual = commandOptionsSchema.safeParse({ name: "test container" });
    assert.strictEqual(actual.success, true);
  });

  it('retrieves container type by ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes/${containerTypeId}`) {
        return containerTypeResponse;
      }

      throw "Invalid request: " + opts.url;
    });

    await command.action(logger, { options: { id: containerTypeId } });
    assert(loggerLogSpy.calledWith(containerTypeResponse));
  });

  it('retrieves the container type by name successfully', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$filter=name eq '${formatting.encodeQueryParameter(containerTypeName)}'`) {
        return {
          value: [
            containerTypeResponse
          ]
        };
      }

      throw "Invalid request: " + opts.url;
    });

    await command.action(logger, { options: { name: containerTypeName } });
    assert(loggerLogSpy.calledWith(containerTypeResponse));
  });

  it('correctly throws error when container type was not found by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$filter=name eq '${formatting.encodeQueryParameter(containerTypeName)}'`) {
        return {
          value: []
        };
      }

      throw "Invalid request: " + opts.url;
    });

    await assert.rejects(command.action(logger, { options: { name: containerTypeName } }),
      new CommandError(`The specified container type '${containerTypeName}' does not exist.`));
  });

  it('correctly prompts when multiple containers found by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes?$filter=name eq '${formatting.encodeQueryParameter(containerTypeName)}'`) {
        return {
          value: [
            containerTypeResponse,
            {
              id: 'a11f96a4-0d60-4cc0-8a82-dcf15591cc02',
              name: containerTypeName
            }
          ]
        };
      }

      throw "Invalid request: " + opts.url;
    });

    const resultsFoundStub = sinon.stub(cli, 'handleMultipleResultsFound').resolves(containerTypeResponse);

    await command.action(logger, { options: { name: containerTypeName } });
    assert(resultsFoundStub.calledOnce);
    assert(loggerLogSpy.calledWith(containerTypeResponse));
  });
});
