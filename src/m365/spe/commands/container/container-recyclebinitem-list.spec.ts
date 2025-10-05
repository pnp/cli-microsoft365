import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { spe } from '../../../../utils/spe.js';
import commands from '../../commands.js';
import command, { options } from './container-recyclebinitem-list.js';

describe(commands.CONTAINER_RECYCLEBINITEM_LIST, () => {
  const containerTypeId = 'dda3cb36-a16a-40b9-8f04-b01e39fc035d';
  const requestResponse = [
    {
      id: 'b!ISJs1WRro0y0EWgkUYcktDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
      displayName: 'Playground container',
      containerTypeId: containerTypeId,
      createdDateTime: '2025-04-15T21:04:25Z',
      settings: {
        isOcrEnabled: true
      }
    },
    {
      id: 'b!3vQnoI2C-UOm3Z_bCtysBbDa0mE8zSlFEqFzqRn70Zwp1CEtDEBZgQICPkRbil_5Z',
      displayName: 'My Application Storage Container',
      containerTypeId: containerTypeId,
      createdDateTime: '2025-04-15T21:51:48Z',
      settings: {
        isOcrEnabled: false
      }
    }
  ];

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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CONTAINER_RECYCLEBINITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName']);
  });

  it('fails validation if both containerTypeId and containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({ containerTypeId: containerTypeId, containerTypeName: 'Container name' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if neither containerTypeId nor containerTypeName options are passed', async () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if containerTypeId is not a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ containerTypeId: 'invalid' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if containerTypeId is a valid GUID', async () => {
    const actual = commandOptionsSchema.safeParse({ containerTypeId: containerTypeId });
    assert.strictEqual(actual.success, true);
  });

  it('correctly outputs a result when using containerTypeId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`) {
        return {
          value: requestResponse
        };
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    await command.action(logger, { options: { containerTypeId: containerTypeId } });
    assert(loggerLogSpy.calledOnceWith(requestResponse));
  });

  it('correctly outputs a result when using containerTypename', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`) {
        return {
          value: requestResponse
        };
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    await command.action(logger, { options: { containerTypeName: 'Container Type Name' } });
    assert(loggerLogSpy.calledOnceWith(requestResponse));
  });

  it('retrieves list of container recycle bin items by using containerTypeId', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`) {
        return {
          value: requestResponse
        };
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    await command.action(logger, { options: { containerTypeId: containerTypeId } });
    assert(getStub.calledOnce);
  });

  it('retrieves list of container recycle bin items by using containerTypeName', async () => {
    const getStub = sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/storage/fileStorage/deletedContainers?$filter=containerTypeId eq ${containerTypeId}`) {
        return {
          value: requestResponse
        };
      }

      throw 'Invalid GET request: ' + opts.url;
    });

    await command.action(logger, { options: { containerTypeName: 'Container Type Name', verbose: true } });
    assert(getStub.calledOnce);
  });

  it('correctly handles unexpected error', async () => {
    const errorMessage = 'Access denied';
    sinon.stub(request, 'get').rejects({
      error: {
        code: 'accessDenied',
        message: errorMessage
      }
    });

    await assert.rejects(command.action(logger, { options: { containerTypeId: containerTypeId } })
      , new CommandError(errorMessage));
  });
});