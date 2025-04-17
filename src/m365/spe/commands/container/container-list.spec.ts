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
import { CommandError } from '../../../../Command.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { cli } from '../../../../cli/cli.js';
import { spe } from '../../../../utils/spe.js';

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

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');

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

    sinon.stub(spe, 'getContainerTypeIdByName').withArgs(adminUrl, 'standard container').resolves('e2756c4d-fa33-4452-9c36-2325686e1082');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      spe.getContainerTypeIdByName
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

    await command.action(logger, { options: { containerTypeId: "e2756c4d-fa33-4452-9c36-2325686e1082", verbose: true } });
    assert(loggerLogSpy.calledWith(containersList));
  });

  it('retrieves list of container type by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/storage/fileStorage/containers?$filter=containerTypeId eq e2756c4d-fa33-4452-9c36-2325686e1082') {
        return { "value": containersList };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { containerTypeName: "standard container", verbose: true } });
    assert(loggerLogSpy.calledWith(containersList));
  });

  it('correctly handles error when retrieving containers', async () => {
    const error = 'An error has occurred';
    sinonUtil.restore(spe.getContainerTypeIdByName);
    sinon.stub(spe, 'getContainerTypeIdByName').rejects(new Error(error));

    await assert.rejects(command.action(logger, {
      options: {
        containerTypeName: "nonexisting container",
        verbose: true
      }
    }), new CommandError('An error has occurred'));
  });
});