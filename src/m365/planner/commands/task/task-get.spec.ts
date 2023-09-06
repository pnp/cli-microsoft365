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
import command from './task-get.js';

describe(commands.TASK_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTaskTitle = 'Task name';
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanTitle = 'Plan title';
  const validRosterId = 'DjL5xiKO10qut8LQgztpKskABWna';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const invalidOwnerGroupId = 'Invalid GUID';

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const multipleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      },
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const singlePlanResponse = {
    "value": [
      {
        "id": validPlanId,
        "title": validPlanTitle
      }
    ]
  };

  const singleBucketByNameResponse = {
    "value": [
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const multipleBucketByNameResponse = {
    "value": [
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      },
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
        "name": validBucketName,
        "id": validBucketId
      }
    ]
  };

  const singleTaskByTitleResponse = {
    "value": [
      {
        "title": validTaskTitle,
        "id": validTaskId
      }
    ]
  };

  const multipleTasksByTitleResponse = {
    "value": [
      {
        "title": validTaskTitle,
        "id": validTaskId
      },
      {
        "title": validTaskTitle,
        "id": validTaskId
      }
    ]
  };

  const taskResponse = {
    "planId": validPlanId,
    "bucketId": validBucketId,
    "title": validTaskTitle,
    "id": validTaskId
  };

  const taskDetailsResponse = {
    "description": "Test",
    "references": {}
  };

  const planResponse = {
    "id": validPlanId,
    "title": validPlanTitle
  };

  const outputResponse = {
    ...taskResponse,
    ...taskDetailsResponse
  };

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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_GET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when bucket name is used without id', async () => {
    const actual = await command.validate({
      options: {
        id: validTaskId,
        bucketName: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is used without bucket id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is used with both bucket id and bucketname', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        bucketName: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used without plan title, plan id, or roster id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used with both plan title, plan id, and roster id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId,
        planTitle: validPlanTitle,
        rosterId: validRosterId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan title is used without owner group name or owner group id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan title is used with both owner group name and owner group id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and plan details are specified', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId,
        planId: validPlanId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when owner group id is not a guid', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validTaskId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when no groups found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return { "value": [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError(`The specified group '${validOwnerGroupName}' does not exist.`));
  });

  it('fails validation when multiple groups found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return multipleGroupResponse;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return { "value": [{ "id": "" }] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`The specified bucket ${validBucketName} does not exist`));
  });

  it('fails validation when multiple buckets found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return multipleBucketByNameResponse;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId
      }
    }), new CommandError("Multiple buckets with name 'Bucket name' found. Found: vncYUXCRBke28qMLB-d4xJcACtNz."));
  });

  it('handles selecting single result when multiple buckets with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return multipleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleBucketByNameResponse.value[0]);

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('fails validation when no tasks found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return { "value": [{ "id": "" }] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }), new CommandError(`The specified task ${validTaskTitle} does not exist`));
  });

  it('fails validation when multiple tasks found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return multipleTasksByTitleResponse;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }), new CommandError("Multiple tasks with title 'Task name' found. Found: 2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2."));
  });

  it('handles selecting single result when multiple tasks with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return multipleTasksByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleTaskByTitleResponse.value[0]);

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly gets task by name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return singleGroupResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly gets task by name with group ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: validOwnerGroupId
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly gets task by task ID', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        id: validTaskId
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly gets task by rosterId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return { "value": [planResponse] };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return taskResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        rosterId: validRosterId
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly handles item not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('The requested item is not found.'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('The requested item is not found.'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
