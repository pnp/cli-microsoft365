import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from "../../../../cli/CommandInfo.js";
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spe } from '../../../../utils/spe.js';
import commands from '../../commands.js';
import command, { options } from './container-add.js';

describe(commands.CONTAINER_ADD, () => {
  const spoAdminUrl = 'https://contoso-admin.sharepoint.com';
  const containerTypeId = 'c6f08d91-77fa-485f-9369-f246ec0fc19c';
  const containerTypeName = 'Container type name';
  const containerName = 'Invoices';

  const requestResponse = {
    id: 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
    displayName: containerName,
    description: 'Description of My Application Storage Container',
    containerTypeId: containerTypeId,
    status: 'inactive',
    createdDateTime: '2025-04-15T13:31:09.62Z',
    lockState: 'unlocked',
    settings: {
      isOcrEnabled: false,
      itemMajorVersionLimit: 500,
      isItemVersioningEnabled: true
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

    sinon.stub(spe, 'getContainerTypeIdByName').resolves(containerTypeId);

    auth.connection.active = true;
    auth.connection.spoUrl = spoAdminUrl.replace('-admin.sharepoint.com', '.sharepoint.com');
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both containerTypeId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeId: containerTypeId, containerTypeName: containerTypeName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither containerTypeId nor containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerTypeId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if containerTypeId is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, containerTypeId: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation if itemMajorVersionLimit is not a positive integer', async () => {
    const actual = commandOptionsSchema.safeParse({ name: containerName, itemMajorVersionLimit: 12.5 });
    assert.strictEqual(actual.success, false);
  });

  it('correctly logs an output', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers') {
        return requestResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId } });
    assert(loggerLogSpy.calledOnceWith(requestResponse));
  });

  it('correctly creates a new container with containerTypeId', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers') {
        return requestResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, description: 'Lorem ipsum', ocrEnabled: true, itemMajorVersionLimit: 250, itemVersioningEnabled: true, containerTypeId: containerTypeId } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: containerName,
      description: 'Lorem ipsum',
      containerTypeId: containerTypeId,
      settings: {
        isOcrEnabled: true,
        itemMajorVersionLimit: 250,
        isItemVersioningEnabled: true
      }
    });
  });

  it('correctly creates a new container with containerTypeName', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers') {
        return requestResponse;
      }

      throw 'Invalid POST request: ' + opts.url;
    });

    await command.action(logger, { options: { name: containerName, description: 'Lorem ipsum', ocrEnabled: true, itemMajorVersionLimit: 250, itemVersioningEnabled: true, containerTypeName: containerTypeName, verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data, {
      displayName: containerName,
      description: 'Lorem ipsum',
      containerTypeId: containerTypeId,
      settings: {
        isOcrEnabled: true,
        itemMajorVersionLimit: 250,
        isItemVersioningEnabled: true
      }
    });
  });

  it('correctly handles error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        code: 'accessDenied',
        message: 'Access denied'
      }
    });

    await assert.rejects(command.action(logger, { options: { name: containerName, containerTypeId: containerTypeId } }),
      new CommandError('Access denied'));
  });
});