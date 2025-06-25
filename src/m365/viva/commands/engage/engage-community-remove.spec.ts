import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import command from './engage-community-remove.js';
import { vivaEngage } from '../../../../utils/vivaEngage.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';

describe(commands.ENGAGE_COMMUNITY_REMOVE, () => {
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9';
  const displayName = 'Software Engineers';
  const entraGroupId = '0bed8b86-5026-4a93-ac7d-56750cc099f1';

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean;
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(async () => {
      promptIssued = true;
      return false;
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ENGAGE_COMMUNITY_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when entraGroupId is specified', async () => {
    const actual = await command.validate({ options: { entraGroupId: entraGroupId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when entraGroupId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { entraGroupId: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('prompts before removing the community when confirm option not passed', async () => {
    await command.action(logger, { options: { id: communityId } });

    assert(promptIssued);
  });

  it('aborts removing the community when prompt not confirmed', async () => {
    const deleteSpy = sinon.stub(request, 'delete').resolves();

    await command.action(logger, { options: { id: communityId } });
    assert(deleteSpy.notCalled);
  });

  it('removes the community specified by id without prompting for confirmation', async () => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { id: communityId, force: true, verbose: true } });
    assert(deleteRequestStub.called);
  });

  it('removes the community specified by displayName while prompting for confirmation', async () => {
    sinon.stub(vivaEngage, 'getCommunityByDisplayName').resolves({ id: communityId });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { displayName: displayName } });
    assert(deleteRequestStub.called);
  });

  it('removes the community specified by Entra group id while prompting for confirmation', async () => {
    sinon.stub(vivaEngage, 'getCommunityByEntraGroupId').resolves({ id: communityId });

    const deleteRequestStub = sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        return;
      }

      throw 'Invalid request';
    });

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, { options: { entraGroupId: entraGroupId } });
    assert(deleteRequestStub.called);
  });

  it('throws an error when the community specified by id cannot be found', async () => {
    const error = {
      error: {
        code: 'notFound',
        message: 'Not found.',
        innerError: {
          date: '2024-08-30T06:25:04',
          'request-id': '186480bb-73a7-4164-8a10-b05f45a94a4f',
          'client-request-id': '186480bb-73a7-4164-8a10-b05f45a94a4f'
        }
      }
    };
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}`) {
        throw error;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { id: communityId, force: true } }),
      new CommandError(error.error.message));
  });
});