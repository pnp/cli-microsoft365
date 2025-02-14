import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import { cli } from '../../../../cli/cli.js';
import request from '../../../../request.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import command from './multitenant-remove.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MULTITENANT_REMOVE, () => {
  const tenantId = "526dcbd1-4f42-469e-be90-ba4a7c0b7802";
  const organization = {
    "id": "526dcbd1-4f42-469e-be90-ba4a7c0b7802"
  };
  const multitenantOrganizationMembers = [
    {
      "tenantId": "526dcbd1-4f42-469e-be90-ba4a7c0b7802"
    },
    {
      "tenantId": "6babcaad-604b-40ac-a9d7-9fd97c0b779f"
    }
  ];
  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.get,
      cli.handleMultipleResultsFound,
      cli.promptForConfirmation,
      global.setTimeout
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MULTITENANT_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the multitenant organization without prompting for confirmation', async () => {
    let i = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id`) {
        return {
          value: [
            organization
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants?$select=tenantId`) {
        if (i++ < 2) {
          return {
            value: multitenantOrganizationMembers
          };
        }
        return {
          value: [
            multitenantOrganizationMembers[0]
          ]
        };
      }

      throw 'Invalid request';
    });
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/${multitenantOrganizationMembers[0].tenantId}`
        || opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/${multitenantOrganizationMembers[1].tenantId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });

    await command.action(logger, { options: { force: true, verbose: true } });
    assert(deleteRequestStub.calledTwice);
  });

  it('removes the multitenant organization while prompting for confirmation', async () => {
    let i = 0;
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id`) {
        return {
          value: [
            organization
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants?$select=tenantId`) {
        if (i++ < 2) {
          return {
            value: multitenantOrganizationMembers
          };
        }
        return {
          value: [
            multitenantOrganizationMembers[0]
          ]
        };
      }

      throw 'Invalid request';
    });
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/${multitenantOrganizationMembers[0].tenantId}`
        || opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/${multitenantOrganizationMembers[1].tenantId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinon.stub(global, 'setTimeout').callsFake((fn) => {
      fn();
      return {} as any;
    });
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: {} });
    assert(deleteRequestStub.calledTwice);
  });

  it('prompts before removing the multitenant organization when prompt option not passed', async () => {
    await command.action(logger, { options: {} });

    assert(promptIssued);
  });

  it('aborts removing the multitenant organization when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: {} });
    assert(deleteSpy.notCalled);
  });

  it('throws an error when one of the tenant cannot be found', async () => {
    const error = {
      error: {
        code: 'Request_ResourceNotFound',
        message: `Resource '${tenantId}' does not exist or one of its queried reference-property objects are not present.`,
        innerError: {
          date: '2024-05-07T06:59:51',
          'request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b',
          'client-request-id': 'b7dee9ee-d85b-4e7a-8686-74852cbfd85b'
        }
      }
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/organization?$select=id`) {
        return {
          value: [
            organization
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants?$select=tenantId`) {
        return {
          value: multitenantOrganizationMembers
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization/tenants/${multitenantOrganizationMembers[1].tenantId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { force: true } }),
      new CommandError(error.error.message));
  });
});