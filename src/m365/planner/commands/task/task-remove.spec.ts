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
import command from './task-remove.js';

describe(commands.TASK_REMOVE, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTaskTitle = 'Task name';
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanTitle = 'Plan name';
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

  const singleTaskByIdResponse = {
    "title": validTaskTitle,
    "id": validTaskId
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

  const planResponse = {
    "id": validPlanId,
    "title": validPlanTitle
  };

  before(() => {
    cli = Cli.getInstance();
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      Cli.prompt,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when title and id is used', async () => {
    const actual = await command.validate({
      options: {
        id: validTaskId,
        title: validTaskTitle
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

  it('fails validation when bucket name is used without plan name, plan id, or roster id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used with both plan name, plan id, and roster id', async () => {
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

  it('fails validation when plan name is used without owner group name or owner group id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan name is used with both owner group name and owner group id', async () => {
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

  it('fails validation when task id is used with bucket id', async () => {
    const actual = await command.validate({
      options: {
        id: validTaskId,
        bucketId: validBucketId
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

  it('validates for a correct input with title', async () => {
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
        ownerGroupName: validOwnerGroupName,
        force: true
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
        ownerGroupName: validOwnerGroupName,
        force: true
      }
    }), new CommandError("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return { "value": [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId,
        force: true
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
        planId: validPlanId,
        force: true
      }
    }), new CommandError("Multiple buckets with name 'Bucket name' found. Found: vncYUXCRBke28qMLB-d4xJcACtNz."));
  });

  it('handles selecting single result when multiple buckets with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        removeRequestIssued = true;
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleBucketByNameResponse.value[0]);

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert(removeRequestIssued);
  });

  it('fails validation when no tasks found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return { "value": [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        force: true
      }
    }), new CommandError(`The specified task ${validTaskTitle} does not exist`));
  });

  it('fails validation when multiple tasks found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return multipleTasksByTitleResponse;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        force: true
      }
    }), new CommandError("Multiple tasks with title 'Task name' found. Found: 2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2."));
  });

  it('handles selecting single result when multiple tasks with the specified name found and cli is set to prompt', async () => {
    let removeRequestIssued = false;

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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return multipleTasksByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        removeRequestIssued = true;
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleTaskByTitleResponse.value[0]);

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert(removeRequestIssued);
  });

  it('prompts before removing the specified task when confirm option not passed with id', async () => {
    await command.action(logger, {
      options: {
        id: validTaskId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified task when confirm option not passed and prompt not confirmed', async () => {
    const postSpy = sinon.spy(request, 'delete');

    await command.action(logger, {
      options: {
        id: validTaskId
      }
    });

    assert(postSpy.notCalled);
  });

  it('correctly deletes task by id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        id: validTaskId,
        force: true
      }
    });
  });

  it('correctly deletes task by title with group id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: validOwnerGroupId
      }
    });
  });

  it('correctly deletes task by title', async () => {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').resolves({ continue: true });

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
  });

  it('correctly deletes task by rosterId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return { "value": [planResponse] };
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return singleBucketByNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return singleTaskByTitleResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return singleTaskByIdResponse;
      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'delete').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return;
      }

      throw 'Invalid Request';
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        rosterId: validRosterId
      }
    });
  });
});
