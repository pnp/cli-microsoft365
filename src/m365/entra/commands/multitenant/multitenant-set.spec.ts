import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import commands from '../../commands.js';
import command, { options } from './multitenant-set.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';

describe(commands.MULTITENANT_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    const commandInfo = cli.getCommandInfo(command);
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
    (command as any).pollingInterval = 0;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.MULTITENANT_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when only displayName is specified', async () => {
    const parseResult = commandOptionsSchema.safeParse({ displayName: 'Contoso organization' });
    assert.strictEqual(parseResult.success, true);
  });

  it('passes validation when only description is specified', async () => {
    const parseResult = commandOptionsSchema.safeParse({ description: 'Contoso and partners' });
    assert.strictEqual(parseResult.success, true);
  });

  it('passes validation when the displayName and description are specified', async () => {
    const parseResult = commandOptionsSchema.safeParse({ displayName: 'Contoso organization', description: 'Contoso and partners' });
    assert.strictEqual(parseResult.success, true);
  });

  it('fails validation when no option is specified', async () => {
    const parseResult = commandOptionsSchema.safeParse({});
    assert.strictEqual(parseResult.success, false);
  });

  it('updates a displayName of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'Contoso organization' }) });
    assert(patchRequestStub.called);
  });

  it('updates a description of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ description: 'Contoso and partners', verbose: true }) });
    assert(patchRequestStub.called);
  });

  it('updates displayName and description of a multitenant organization', async () => {
    const patchRequestStub = sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/tenantRelationships/multiTenantOrganization') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'Contoso organization', description: 'Contoso and partners' }) });
    assert(patchRequestStub.called);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'patch').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({ displayName: 'test' }) }), new CommandError('Invalid request'));
  });
});