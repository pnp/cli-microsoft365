import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./task-get');

describe(commands.TASK_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTaskTitle = 'Task name';
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanName = 'Plan name';
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
        "title": validPlanName
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

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
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
      request.get,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets();
    assert.deepStrictEqual(optionSets, [['id', 'title']]);
  });

  it('fails validation when title is used without bucket id', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is used with both bucket id and bucketname', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketId: validBucketId,
        bucketName: validBucketName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used without plan name or plan id', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucket name is used with both plan name and plan id', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        bucketName: validBucketName,
        planId: validPlanId,
        planName: validPlanName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan name is used without owner group name or owner group id', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan name is used with both owner group name and owner group id', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName,
        ownerGroupId: validOwnerGroupId
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when owner group id is not a guid', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupId: invalidOwnerGroupId
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', () => {
    const actual = command.validate({
      options: {
        id: validTaskId
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', () => {
    const actual = command.validate({
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation when using app only access token', (done) => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    command.action(logger, {
      options: {
        id: validTaskId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('This command does not support application permissions.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no groups found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'&$select=id`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified ownerGroup ${validOwnerGroupName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple groups found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'&$select=id`) {
        return Promise.resolve(multipleGroupResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple ownerGroups with name ${validOwnerGroupName} found: ${multipleGroupResponse.value.map(x => x.id)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no buckets found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve({ "value": [{ "id": "" }] });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified bucket ${validBucketName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple buckets found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets?$select=id,name`) {
        return Promise.resolve(multipleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planId: validPlanId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple buckets with name ${validBucketName} found: ${multipleBucketByNameResponse.value.map(x => x.id)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no tasks found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve({ "value": [{ "id": "" }] });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified task ${validTaskTitle} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple tasks found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}/tasks?$select=id,title`) {
        return Promise.resolve(multipleTasksByTitleResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketId: validBucketId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple tasks with title ${validTaskTitle} found: ${multipleTasksByTitleResponse.value.map(x => x.id)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly gets task by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'&$select=id`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly gets task by name with group ID', (done) => {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(validTaskId)}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        title: validTaskTitle,
        bucketName: validBucketName,
        planName: validPlanName,
        ownerGroupId: validOwnerGroupId
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('successfully handles item found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/01gzSlKkIUSUl6DF_EilrmQAKDhh`) {
        return Promise.resolve({
          "createdBy": {
            "user": {
              "id": "6463a5ce-2119-4198-9f2a-628761df4a62"
            }
          },
          "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
          "bucketId": "gcrYAaAkgU2EQUvpkNNXLGQAGTtu",
          "title": "title-value",
          "orderHint": "9223370609546166567W",
          "assigneePriority": "90057581\"",
          "createdDateTime": "2015-03-25T18:36:49.2407981Z",
          "assignments": {
            "fbab97d0-4932-4511-b675-204639209557": {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              "assignedBy": {
                "user": {
                  "id": "1e9955d2-6acd-45bf-86d3-b546fdc795eb"
                }
              },
              "assignedDateTime": "2015-03-25T18:38:21.956Z",
              "orderHint": "RWk1"
            }
          },
          "priority": 5,
          "id": "01gzSlKkIUSUl6DF_EilrmQAKDhh"
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        id: '01gzSlKkIUSUl6DF_EilrmQAKDhh', debug: true
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify({
          "createdBy": {
            "user": {
              "id": "6463a5ce-2119-4198-9f2a-628761df4a62"
            }
          },
          "planId": "xqQg5FS2LkCp935s-FIFm2QAFkHM",
          "bucketId": "gcrYAaAkgU2EQUvpkNNXLGQAGTtu",
          "title": "title-value",
          "orderHint": "9223370609546166567W",
          "assigneePriority": "90057581\"",
          "createdDateTime": "2015-03-25T18:36:49.2407981Z",
          "assignments": {
            "fbab97d0-4932-4511-b675-204639209557": {
              "@odata.type": "#microsoft.graph.plannerAssignment",
              "assignedBy": {
                "user": {
                  "id": "1e9955d2-6acd-45bf-86d3-b546fdc795eb"
                }
              },
              "assignedDateTime": "2015-03-25T18:38:21.956Z",
              "orderHint": "RWk1"
            }
          },
          "priority": 5,
          "id": "01gzSlKkIUSUl6DF_EilrmQAKDhh"
        });
        assert.strictEqual(actual, expected);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles item not found', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('The requested item is not found.'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The requested item is not found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => Promise.reject('An error has occurred'));

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});