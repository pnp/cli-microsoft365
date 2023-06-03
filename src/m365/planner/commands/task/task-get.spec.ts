import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./task-get');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
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
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return Promise.resolve(multipleGroupResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError(`Multiple groups with name '${validOwnerGroupName}' found: ${multipleGroupResponse.value.map(x => x.id)}.`));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve({ "value": [{ "id": "" }] });
      }

      return Promise.reject('Invalid Request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(multipleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`Multiple buckets with name ${validBucketName} found: ${multipleBucketByNameResponse.value.map(x => x.id)}`));
  });

  it('fails validation when no tasks found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve({ "value": [{ "id": "" }] });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }), new CommandError(`The specified task ${validTaskTitle} does not exist`));
  });

  it('fails validation when multiple tasks found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve(multipleTasksByTitleResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }), new CommandError(`Multiple tasks with title ${validTaskTitle} found: ${multipleTasksByTitleResponse.value.map(x => x.id)}`));
  });

  it('correctly gets task by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return Promise.resolve(singleGroupResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve(singleTaskByTitleResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return Promise.resolve(taskDetailsResponse);
      }

      return Promise.reject('Invalid Request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve(singleTaskByTitleResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return Promise.resolve(taskDetailsResponse);
      }

      return Promise.reject('Invalid Request');
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return Promise.resolve(taskDetailsResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        id: validTaskId
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly gets task by rosterId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${validRosterId}/plans`) {
        return Promise.resolve({ "value": [planResponse] });
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve(singleTaskByTitleResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return Promise.resolve(taskDetailsResponse);
      }

      return Promise.reject('Invalid Request');
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
    sinon.stub(request, 'get').callsFake(() => Promise.reject('The requested item is not found.'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('The requested item is not found.'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
