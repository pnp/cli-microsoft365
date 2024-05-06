import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { Logger } from '../../../../cli/Logger.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import command from './multitenant-get.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MULTITENANT_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const response = {
    "id": "ab217953-e37f-4691-97b8-dbb8a0a3bcaf",
    "createdDateTime": "2024-05-05T05:05:05",
    "state": "active",
    "displayName": "Contoso organization",
    "description": "Contoso and partners"
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MULTITENANT_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves information about the multitenant organization', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return response;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { } });
    assert(loggerLogSpy.calledOnceWithExactly(response));
  });

  it('handles random API error', async () => {
    const errorMessage = 'Something went wrong';
    sinon.stub(request, 'get').rejects(new Error(errorMessage));

    await assert.rejects(command.action(logger, { options: {} }), new CommandError(errorMessage));
  });
});