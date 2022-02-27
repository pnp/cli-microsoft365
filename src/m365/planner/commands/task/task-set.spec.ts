import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./task-set');

describe(commands.TASK_SET, () => {
  const taskResponse = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {}
  };

  const taskResponseWithDetails = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {},
    "description": "My Task Description",
    "references": {},
    "checklist": {}
  };

  const taskResponseWithAssignments = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "Z-RLQGfppU6H3663DBzfs5gAMD3o",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {
      "category1": true,
      "category2": true
    },
    "assignments": {
      "949b16c1-a032-453e-a8ae-89a52bfc1d8a": {
        "assignedDateTime": "2021-12-12T19:39:28.0946158Z",
        "orderHint": "8585622689774142174P}",
        "assignedBy": {
          "user": {
            "displayName": null,
            "id": "dd8b99a7-77c6-4238-a609-396d27844921"
          }
        }
      }
    }
  };

  const groupByDisplayNameResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": "0d0402ee-970f-4951-90b5-2f24519d2e40"
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    (command as any).planId = undefined;
    (command as any).bucketId = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post,
      request.patch
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when both bucketId and bucketName are specified', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        bucketName: 'My Bucket'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when bucketName is specified but not planId or planName', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when bucketName is specified but both planId and planName are specified', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        planName: 'My Planner'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified without ownerGroupId or ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planName: 'My Planner Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified with both ownerGroupId and ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planName: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Bucket',
        planName: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if startDateTime contains invalid format.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        startDateTime: '2021-99-99'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if dueDateTime contains invalid format.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        dueDateTime: '2021-99-99'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if percentComplete contains invalid format.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        percentComplete: 'Not A Number'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if percentComplete is not between 0 and 100.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        percentComplete: 599
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if assignedToUserIds contains invalid guid.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-test"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both assignedToUserIds and assignedToUserNames are specified', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-eb6030e1a09e",
        assignedToUserNames: "Allan.Carroll@contoso.onmicrosoft.com,Ida.Stevens@contoso.onmicrosoft.com"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if incorrect appliedCategory is specified.', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        appliedCategories: "category1,category9"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid options specified', (done) => {
    const actual = command.validate({
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        title: 'My Planner Task'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('correctly updates planner task with title', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      title: 'My Planner Task'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task  to bucket with bucketName, planName, and ownerGroupName', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${encodeURIComponent('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return Promise.resolve({
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')&$select=id,title`) {
        return Promise.resolve({
          value: [
            {
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'&$select=id`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planName: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task  to bucket with bucketName, planName, and ownerGroupId', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${encodeURIComponent('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return Promise.resolve({
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')&$select=id,title`) {
        return Promise.resolve({
          value: [
            {
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planName: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task  to bucket with bucketName, planId', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${encodeURIComponent('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return Promise.resolve({
          value: [
            {
              "name": "My Planner Bucket",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task with assignedToUserIds', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponseWithAssignments);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserIds: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponseWithAssignments));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task with assignedToUserNames', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponseWithAssignments);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${encodeURIComponent('user@contoso.onmicrosoft.com')}'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserNames: 'user@contoso.onmicrosoft.com'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponseWithAssignments));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task with description', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return Promise.resolve(taskResponseWithDetails);
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        return Promise.resolve({
          "description": "My Task Description",
          "references": {},
          "checklist": {}
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponseWithDetails);
      }

      return Promise.reject('Invalid Request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      description: 'My Task Description'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponseWithDetails));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly updates planner task with appliedCategories, bucketId, startDateTime, dueDateTime, percentComplete, assigneePriority, and orderHint', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      appliedCategories: 'category1,category2',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      startDateTime: '2014-01-01T00:00:00Z',
      dueDateTime: '2023-01-01T00:00:00Z',
      percentComplete: '50',
      assigneePriority: ' !',
      orderHint: ' !'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when no bucket is found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${encodeURIComponent('8QZEH7b3wkS_bGQobscsM5gADCBb')}/buckets?$select=id,name`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      bucketName: 'My Planner Bucket',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "The specified bucket does not exist");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return Promise.resolve({
          value: []
        });
      }

      return Promise.reject('Invalid request');
    });

    const options: any = {
      id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
      assignedToUserNames: 'user@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com'
    };

    command.action(logger, { options: options } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, "Cannot proceed with planner task update. The following users provided are invalid : user2@contoso.onmicrosoft.com");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when ownerGroupName not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Planner Bucket',
        planName: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified owner group does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when planName not found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'&$select=id`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')&$select=id,title`) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        bucketName: 'My Planner Bucket',
        planName: 'foo',
        ownerGroupName: 'My Planner Group'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified plan does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when task endpoint fails', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        return Promise.resolve(undefined);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve(undefined);
      }

      return Promise.reject('Invalid request');
    });


    command.action(logger, {
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        title: 'My Planner Task'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Error fetching task`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when task details endpoint fails', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}`) {
        return Promise.resolve(taskResponse);
      }

      return Promise.reject('Invalid Request');
    });

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}/details`) {
        return Promise.resolve(undefined);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent('Z-RLQGfppU6H3663DBzfs5gAMD3o')}` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return Promise.resolve({
          "@odata.etag": "TestEtag"
        });
      }

      return Promise.reject('Invalid request');
    });


    command.action(logger, {
      options: {
        id: 'Z-RLQGfppU6H3663DBzfs5gAMD3o',
        description: 'My Task Description'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Error fetching task details`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles random API error', (done) => {
    Utils.restore(request.get);
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