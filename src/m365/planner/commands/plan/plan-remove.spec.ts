import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
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
import command from './plan-remove.js';

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
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = Cli.getCommandInfo(command);
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
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PLAN_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id and ownerGroupId is specified', async () => {
    const actual = await command.validate({
      options: {
        id: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified with both ownerGroupName and ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified without ownerGroupName or ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified with invalid ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validPlanId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with title', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified plan when confirm option not passed with id', async () => {
    await command.action(logger, {
      options: {
        id: validPlanId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified plan when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        id: validPlanId
      }
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
      options: {
        id: validPlanId,
        force: true
      }
    });
  });

  it('Correctly deletes plan by title', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
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
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
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
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        title: validPlanTitle,
        ownerGroupId: validOwnerGroupId
      }
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').resolves(singlePlanResponse);
    sinon.stub(request, 'delete').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        id: validPlanId,
        force: true
      }
    }), new CommandError("An error has occurred"));
  });
});
