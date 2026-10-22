import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './container-permission-list.js';
import { z } from 'zod';
import { spe } from '../../../../utils/spe.js';
import { odata } from '../../../../utils/odata.js';
import { cli } from '../../../../cli/cli.js';

describe(commands.CONTAINER_PERMISSION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let schema: z.ZodTypeAny;

  const containerId = "b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z";
  const containerName = 'My Application Storage Container';
  const containerTypeId = 'b2e2cef4-9ac1-4b3b-b4a5-2a2e3a2e2a2e';
  const containerPermissionResponse = {
    "value": [
      {
        "id": "X2k6MCMuZnxtZW1iZXJzaGlwfGRlYnJhYkBuYWNoYW4zNjUub25taWNyb3NvZnQuY29t",
        "roles": [
          "owner"
        ],
        "grantedToV2": {
          "user": {
            "displayName": "Debra Berger",
            "email": "debra@contoso.onmicrosoft.com",
            "userPrincipalName": "debra@contoso.onmicrosoft.com"
          }
        }
      },
      {
        "id": "X2k6MCMuZnxtZW1iZXJzaGlwfGFkbWluQG5hY2hhbjM2NS5vbm1pY3Jvc29mdC5jb20",
        "roles": [
          "reader"
        ],
        "grantedToV2": {
          "user": {
            "displayName": "John Doe",
            "email": "john@contoso.onmicrosoft.com",
            "userPrincipalName": "john@contoso.onmicrosoft.com"
          }
        }
      }
    ]
  };

  const textOutput = [
    {
      "id": "X2k6MCMuZnxtZW1iZXJzaGlwfGRlYnJhYkBuYWNoYW4zNjUub25taWNyb3NvZnQuY29t",
      "roles": "owner",
      "userPrincipalName": "debra@contoso.onmicrosoft.com"
    },
    {
      "id": "X2k6MCMuZnxtZW1iZXJzaGlwfGFkbWluQG5hY2hhbjM2NS5vbm1pY3Jvc29mdC5jb20",
      "roles": "reader",
      "userPrincipalName": "john@contoso.onmicrosoft.com"
    }
  ];

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    auth.connection.active = true;
    schema = command.getSchemaToParse()!;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      spe.getContainerIdByName,
      spe.getContainerTypeIdByName,
      odata.getAllItems,
      cli.handleMultipleResultsFound
    ]);
    loggerLogSpy.restore();
    loggerLogToStderrSpy.restore();
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_PERMISSION_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'userPrincipalName', 'roles']);
  });

  it('correctly lists permissions of a SharePoint Embedded Container', async () => {
    sinon.stub(odata, 'getAllItems').resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerId: containerId,
        debug: true
      }
    });

    assert(loggerLogSpy.calledWith(containerPermissionResponse.value));
  });

  it('correctly lists permissions of a SharePoint Embedded Container (TEXT)', async () => {
    sinon.stub(odata, 'getAllItems').resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerId: containerId,
        debug: true,
        output: 'text'
      }
    });

    assert(loggerLogSpy.calledWith(textOutput));
  });

  it('correctly lists permissions of a SharePoint Embedded Container by name', async () => {
    sinon.stub(odata, 'getAllItems').onFirstCall().resolves([
      {
        id: containerId,
        displayName: containerName
      }
    ]).onSecondCall().resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerName,
        containerTypeId,
        debug: true
      }
    });

    assert(loggerLogSpy.calledWith(containerPermissionResponse.value));
  });

  it('logs progress when resolving container id by name in verbose mode', async () => {
    sinon.stub(odata, 'getAllItems').onFirstCall().resolves([
      {
        id: containerId,
        displayName: containerName
      }
    ]).onSecondCall().resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerName,
        containerTypeId,
        verbose: true
      }
    });

    assert(loggerLogToStderrSpy.calledWith(`Resolving container id from name '${containerName}'...`));
  });

  it('fails when container with specified name does not exist', async () => {
    sinon.stub(odata, 'getAllItems').resolves([]);

    await assert.rejects(
      command.action(logger, {
        options: {
          containerName,
          containerTypeId
        }
      }),
      new CommandError(`The specified container '${containerName}' does not exist.`)
    );
  });

  it('handles multiple containers with same name when resolving id', async () => {
    sinon.stub(odata, 'getAllItems').onFirstCall().resolves([
      {
        id: '1',
        displayName: containerName
      },
      {
        id: containerId,
        displayName: containerName
      }
    ]).onSecondCall().resolves(containerPermissionResponse.value);
    sinon.stub(cli, 'handleMultipleResultsFound').resolves({
      id: containerId
    });

    await command.action(logger, {
      options: {
        containerName,
        containerTypeId
      }
    });

    assert(loggerLogSpy.calledWith(containerPermissionResponse.value));
  });

  it('rethrows unexpected errors when resolving container id by name', async () => {
    sinon.stub(odata, 'getAllItems').rejects({
      error: {
        'odata.error': {
          message: {
            value: 'unexpected error'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, {
      options: {
        containerName,
        containerTypeId
      }
    }), new CommandError('unexpected error'));
  });

  it('rethrows CommandError thrown during command execution', async () => {
    sinon.stub(odata, 'getAllItems').rejects(new CommandError('command error'));

    await assert.rejects(command.action(logger, {
      options: {
        containerId
      }
    }), new CommandError('command error'));
  });

  it('correctly handles error when SharePoint Embedded Container is not found', async () => {
    sinon.stub(odata, 'getAllItems').rejects({
      error: { 'odata.error': { message: { value: 'Item Not Found.' } } }
    });

    await assert.rejects(command.action(logger, { options: { containerId: containerId } } as any),
      new CommandError('Item Not Found.'));
  });

  it('correctly handles error when retrieving permissions of a SharePoint Embedded Container', async () => {
    const error = 'An error has occurred';
    sinon.stub(odata, 'getAllItems').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        containerId: containerId
      }
    }), new CommandError(error));
  });

  it('fails validation when neither containerId nor containerName is specified', () => {
    const result = schema.safeParse({});
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('Specify either id or name')));
  });

  it('fails validation when both containerId and containerName are specified', () => {
    const result = schema.safeParse({
      containerId,
      containerName
    });
    assert.strictEqual(result.success, false);
    assert(result.error?.issues.some(issue => issue.message.includes('Specify either id or name')));
  });

  it('passes validation when only containerId is specified', () => {
    const result = schema.safeParse({ containerId });
    assert.strictEqual(result.success, true);
  });

  it('passes validation when containerName and containerTypeId are specified', () => {
    const result = schema.safeParse({ containerName, containerTypeId });
    assert.strictEqual(result.success, true);
  });

  it('correctly lists permissions of a SharePoint Embedded Container by containerTypeName', async () => {
    const containerTypeName = 'My Container Type';
    sinon.stub(spe, 'getContainerTypeIdByName').resolves(containerTypeId);
    sinon.stub(odata, 'getAllItems').onFirstCall().resolves([
      {
        id: containerId,
        displayName: containerName
      }
    ]).onSecondCall().resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerName,
        containerTypeName
      }
    });

    assert(loggerLogSpy.calledWith(containerPermissionResponse.value));
  });

  it('logs progress when getting container type by name in verbose mode', async () => {
    const containerTypeName = 'My Container Type';
    sinon.stub(spe, 'getContainerTypeIdByName').resolves(containerTypeId);
    sinon.stub(odata, 'getAllItems').onFirstCall().resolves([
      {
        id: containerId,
        displayName: containerName
      }
    ]).onSecondCall().resolves(containerPermissionResponse.value);

    await command.action(logger, {
      options: {
        containerName,
        containerTypeName,
        verbose: true
      }
    });

    assert(loggerLogToStderrSpy.calledWith(`Getting container type with name '${containerTypeName}'...`));
  });
});
