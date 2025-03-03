import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import commands from '../../commands.js';
import command from './multitenant-add.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MULTITENANT_ADD, () => {
  const multitenantOrganizationShortReponse = {
    "id": "ab217953-e37f-4691-97b8-dbb8a0a3bcaf",
    "createdDateTime": "2024-05-03T06:44:57Z",
    "state": "active",
    "displayName": "Contoso organization"
  };

  const multitenantOrganizationReponse = {
    "id": "ab217953-e37f-4691-97b8-dbb8a0a3bcaf",
    "createdDateTime": "2024-05-03T06:44:57Z",
    "state": "active",
    "displayName": "Contoso organization",
    "description": "Contoso and partners"
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.put
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MULTITENANT_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when only displayName is specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Contoso organization' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when the displayName and description are specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Contoso organization', description: 'Contoso and partners' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('creates a multitenant organization with a displayName only', async () => {
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return multitenantOrganizationShortReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Contoso organization', verbose: true } });
    assert(loggerLogSpy.calledOnceWithExactly(multitenantOrganizationShortReponse));
  });

  it('creates a multitenant organization with displayName and description', async () => {
    sinon.stub(request, 'put').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return multitenantOrganizationReponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Contoso organization', description: 'Contoso and partners' } });
    assert(loggerLogSpy.calledOnceWithExactly(multitenantOrganizationReponse));
  });

  it('correctly handles random API OData error', async () => {
    sinon.stub(request, 'put').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Invalid request'));
  });

  it('correctly handles API OData error when the multitenant organization already exist', async () => {
    sinon.stub(request, 'put').rejects({
      error: {
        code: 'Request_BadRequest',
        message: 'Method not supported for update operation.',
        innerError: {
          date: '2024-05-05T05:05:05',
          'request-id': '563f55fe-ea32-48b7-bee4-c9f3ca3a4ecf',
          'client-request-id': '563f55fe-ea32-48b7-bee4-c9f3ca3a4ecf'
        }
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Method not supported for update operation.'));
  });
});