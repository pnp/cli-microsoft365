import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './plan-get.js';

describe(commands.PLAN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTitle = 'Plan name';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const validRosterId = 'FeMZFDoK8k2oWmuGE-XFHZcAEwtn';
  const invalidOwnerGroupId = 'Invalid GUID';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const planResponse = {
    "id": validId,
    "title": validTitle
  };

  const planDetailsResponse = {
    "sharedWith": {},
    "categoryDescriptions": {}
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
  };

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PLAN_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified', async () => {
    const actual = await command.validate({
      options: {
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when title and valid ownerGroupId specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when title and valid ownerGroupName specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when rosterId specified', async () => {
    const actual = await command.validate({
      options: {
        rosterId: validRosterId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly get planner plan with given id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    await command.action(logger, {
      options: {
        id: validId
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given title and ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return { "value": [planResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await command.action(logger, { options: options });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given title and ownerGroupName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return singleGroupResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return { "value": [planResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      title: validTitle,
      ownerGroupName: validOwnerGroupName
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given rosterId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return { "value": [planResponse] };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return planDetailsResponse;
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      rosterId: validRosterId
    };

    await command.action(logger, { options: options });
    assert(loggerLogSpy.calledWith(outputResponse));
  });


  it('correctly handles no plan found with given ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return { "value": [] };
      }

      throw `Invalid request ${opts.url}`;
    });

    const options: any = {
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await assert.rejects(command.action(logger, { options: options } as any));
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').rejects(new Error(`Planner plan with id '${validId}' was not found.`));

    await assert.rejects(command.action(logger, { options: { id: validId } }), new CommandError(`Planner plan with id '${validId}' was not found.`));
  });
});
