import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './plan-remove.js';

describe(commands.PLAN_REMOVE, () => {
  const validPlanTitle = 'My Plan';
  const validPlanId = 'opb7bchfZUiFbVWEPL7jPGUABW7f';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const validOwnerGroupName = 'HR';

  const singlePlanResponse = {
    '@odata.etag': 'abcdef',
    title: validPlanTitle,
    owner: validOwnerGroupId,
    id: validPlanId
  };

  const singleGroupsResponse = {
    value: [
      {
        id: validOwnerGroupId,
        displayName: validOwnerGroupName
      }
    ]
  };

  const singlePlansResponse = {
    value: [
      {
        '@odata.etag': 'abcdef',
        id: validPlanId,
        title: validPlanTitle,
        owner: validOwnerGroupId
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let promptIssued: boolean = false;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PLAN_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id and ownerGroupId is specified', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validPlanId,
      ownerGroupId: validOwnerGroupId
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when title is specified with both ownerGroupName and ownerGroupId', () => {
    const actual = commandOptionsSchema.safeParse({
      title: validPlanTitle,
      ownerGroupId: validOwnerGroupId,
      ownerGroupName: validOwnerGroupName
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when title is specified without ownerGroupName or ownerGroupId', () => {
    const actual = commandOptionsSchema.safeParse({
      title: validPlanTitle
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation when title is specified with invalid ownerGroupId', () => {
    const actual = commandOptionsSchema.safeParse({
      title: validPlanTitle,
      ownerGroupId: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validPlanId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('validates for a correct input with id', () => {
    const actual = commandOptionsSchema.safeParse({
      id: validPlanId
    });
    assert.strictEqual(actual.success, true);
  });

  it('validates for a correct input with title', () => {
    const actual = commandOptionsSchema.safeParse({
      title: validPlanTitle,
      ownerGroupName: validOwnerGroupName
    });
    assert.strictEqual(actual.success, true);
  });

  it('prompts before removing the specified plan when force option not passed with id', async () => {
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validPlanId
      })
    });

    assert(promptIssued);
  });

  it('aborts removing the specified plan when force option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validPlanId
      })
    });
    assert(deleteSpy.notCalled);
  });

  it('Correctly deletes plan by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return singlePlanResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validPlanId,
        force: true
      })
    });
  });

  it('Correctly deletes plan by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'&$select=id`) {
        return singleGroupsResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlansResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      })
    });
  });

  it('Correctly deletes plan by title with group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlansResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return;
      }

      throw 'Invalid request';
    });
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(true);

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        title: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        verbose: true
      })
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').resolves(singlePlanResponse);
    sinon.stub(request, 'delete').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        id: validPlanId,
        force: true
      })
    }), new CommandError('An error has occurred'));
  });
});
