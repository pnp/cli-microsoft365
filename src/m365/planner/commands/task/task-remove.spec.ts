import * as assert from 'assert';
import * as os from 'os';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-remove');

describe(commands.TASK_REMOVE, () => {
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

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'title']]);
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

  it('fails validation when bucket name is used without plan name or plan id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used with both plan name and plan id', async () => {
    const actual = await command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId,
        planTitle: validPlanTitle
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        confirm: true
      }
    }), new CommandError(`The specified group '${validOwnerGroupName}' does not exist.`));
  });

  it('fails validation when multiple groups found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve(multipleGroupResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        confirm: true
      }
    }), new CommandError(`Multiple groups with name '${validOwnerGroupName}' found: ${multipleGroupResponse.value.map(x => x.id)}.`));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId,
        confirm: true
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
        planId: validPlanId,
        confirm: true
      }
    }), new CommandError(`Multiple buckets with name ${validBucketName} found: Please disambiguate:${os.EOL}${multipleBucketByNameResponse.value.map(f => `- ${f.id}`).join(os.EOL)}`));
  });

  it('fails validation when no tasks found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        confirm: true
      }
    }), new CommandError(`The specified task ${validTaskTitle} does not exist`));
  });

  it('fails validation when multiple tasks found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return Promise.resolve(multipleTasksByTitleResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        confirm: true
      }
    }), new CommandError(`Multiple tasks with title ${validTaskTitle} found: Please disambiguate: ${os.EOL}${multipleTasksByTitleResponse.value.map(f => `- ${f.id}`).join(os.EOL)}`));
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve(singleTaskByIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        id: validTaskId,
        confirm: true
      }
    });
  });

  it('correctly deletes task by title with group id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return Promise.resolve(singleTaskByTitleResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve(singleTaskByIdResponse);
      }
      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve(singleGroupResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=title,id`) {
        return Promise.resolve(singleTaskByTitleResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve(singleTaskByIdResponse);
      }
      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${validTaskId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});