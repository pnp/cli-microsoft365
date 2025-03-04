import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './container-permission-list.js';
import { formatting } from '../../../../utils/formatting.js';

describe(commands.CONTAINER_PERMISSION_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const containerId = "b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z";
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

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    auth.connection.active = true;
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/containers/${formatting.encodeQueryParameter(containerId)}/permissions`) {
        return containerPermissionResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        containerId: containerId,
        debug: true
      }
    });

    assert(loggerLogSpy.calledWith(
      [
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
      ]
    ));
  });

  it('correctly handles error when SharePoint Embedded Container is not found', async () => {
    sinon.stub(request, 'get').rejects({
      error: { 'odata.error': { message: { value: 'Item Not Found.' } } }
    });

    await assert.rejects(command.action(logger, { options: { containerId: containerId } } as any),
      new CommandError('Item Not Found.'));
  });

  it('correctly handles error when retrieving permissions of a SharePoint Embedded Container', async () => {
    const error = 'An error has occurred';
    sinon.stub(request, 'get').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        containerId: containerId
      }
    }), new CommandError(error));
  });
});