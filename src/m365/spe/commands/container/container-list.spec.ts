import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './container-list.js';
import { spo } from '../../../../utils/spo.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.CONTAINER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const adminUrl = 'https://contoso-admin.sharepoint.com';
  const containersList = [{
    "id": "b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z",
    "displayName": "My File Storage Container",
    "containerTypeId": "e2756c4d-fa33-4452-9c36-2325686e1082",
    "createdDateTime": "2021-11-24T15:41:52.347Z"
  },
  {
    "id": "b!NdyMBAJ1FEWHB2hEx0DND2dYRB9gz4JOl4rzl7-DuyPG3Fidzm5TTKkyZW2beare",
    "displayName": "Trial Container",
    "containerTypeId": "e2756c4d-fa33-4452-9c36-2325686e1082",
    "createdDateTime": "2021-11-24T15:41:52.347Z"
  }];

  const containerTypedata = [{
    "AzureSubscriptionId": "/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)",
    "ContainerTypeId": "/Guid(e2756c4d-fa33-4452-9c36-2325686e1082)",
    "CreationDate": "3/11/2024 2:38:56 PM",
    "DisplayName": "standard container",
    "ExpiryDate": "3/11/2028 2:38:56 PM",
    "IsBillingProfileRequired": true,
    "OwningAppId": "/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)",
    "OwningTenantId": "/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)",
    "Region": "West Europe",
    "ResourceGroup": "Standard group",
    "SPContainerTypeBillingClassification": "Standard"
  },
  {
    "AzureSubscriptionId": "/Guid(f08575e2-36c4-407f-a891-eabae23f66bc)",
    "ContainerTypeId": "/Guid(e2756c4d-fa33-4452-9c36-2325686e1082)",
    "CreationDate": "3/11/2024 2:38:56 PM",
    "DisplayName": "trial container",
    "ExpiryDate": "3/11/2028 2:38:56 PM",
    "IsBillingProfileRequired": true,
    "OwningAppId": "/Guid(1b3b8660-9a44-4a7c-9c02-657f3ff5d5ac)",
    "OwningTenantId": "/Guid(e1dd4023-a656-480a-8a0e-c1b1eec51e1d)",
    "Region": "West Europe",
    "ResourceGroup": "Standard group",
    "SPContainerTypeBillingClassification": "Standard"
  }];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(spo, 'getSpoAdminUrl').resolves(adminUrl);
    sinon.stub(spo, 'ensureFormDigest').resolves({ FormDigestValue: 'abc', FormDigestTimeoutSeconds: 1800, FormDigestExpiresAt: new Date(), WebFullUrl: 'https://contoso.sharepoint.com' });
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spo.getSpoAdminUrl,
      spo.getAllContainerTypes
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'containerTypeId', 'createdDateTime']);
  });

  it('fails validation if the containerTypeId is not a valid guid', async () => {
    const actual = await command.validate({ options: { containerTypeId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid containerTypeId is specified', async () => {
    const actual = await command.validate({ options: { containerTypeId: "e2756c4d-fa33-4452-9c36-2325686e1082" } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('retrieves list of container type by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq e2756c4d-fa33-4452-9c36-2325686e1082') {
        return { "value": containersList };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { containerTypeId: "e2756c4d-fa33-4452-9c36-2325686e1082", debug: true } });
    assert(loggerLogSpy.calledWith(containersList));
  });

  it('retrieves list of container type by name', async () => {
    sinon.stub(spo, 'getAllContainerTypes').resolves(containerTypedata);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq e2756c4d-fa33-4452-9c36-2325686e1082') {
        return { "value": containersList };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { containerTypeName: "standard container", debug: true } });
    assert(loggerLogSpy.calledWith(containersList));
  });

  it('throws an error when service principal is not found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq e2756c4d-fa33-4452-9c36-2325686e1086') {
        return [];
      }

      throw 'Invalid request';
    });

    sinon.stub(spo, 'getAllContainerTypes').resolves(containerTypedata);

    await assert.rejects(command.action(logger, { options: { containerTypeName: "nonexisting container", debug: true } }),
      new CommandError(`Container type with name nonexisting container not found`));
  });

  it('correctly handles error when retrieving containers', async () => {
    const error = 'An error has occurred';
    sinon.stub(spo, 'getAllContainerTypes').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        debug: true
      }
    }), new CommandError('An error has occurred'));
  });
});