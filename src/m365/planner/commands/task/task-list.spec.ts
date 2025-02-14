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
import command from './task-list.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.TASK_LIST, () => {
  const taskListResponseValue = [
    {
      "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc=\"",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "bucketId": "FtzysDykv0-9s9toWiZhdskAD67z",
      "title": "Bucket Task 1",
      "orderHint": "8585760017701920008P'",
      "assigneePriority": "",
      "percentComplete": 0,
      "startDateTime": null,
      "createdDateTime": "2021-07-06T20:59:35.4105517Z",
      "dueDateTime": null,
      "hasDescription": false,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "id": "KvamtRjaPkmPVy1rEA1r2skAOxcA",
      "createdBy": {
        "user": {
          "displayName": null,
          "id": "73829096-6f0a-4745-8f72-12a17bacadea"
        }
      },
      "appliedCategories": {},
      "assignments": {}
    },
    {
      "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc=\"",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "bucketId": "FtzysDykv0-9s9toWiZhdskAD67z",
      "title": "Bucket Task 2",
      "orderHint": "8585763504689506592PK",
      "assigneePriority": "8585763504089037251",
      "percentComplete": 0,
      "startDateTime": null,
      "createdDateTime": "2021-07-02T20:07:56.5738556Z",
      "dueDateTime": null,
      "hasDescription": false,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "id": "BNWGt05mFUq1VV-cdK00aMkAH5nT",
      "createdBy": {
        "user": {
          "displayName": null,
          "id": "73829096-6f0a-4745-8f72-12a17bacadea"
        }
      },
      "appliedCategories": {},
      "assignments": {}
    }
  ];

  const taskListResponseBetaValue = [
    {
      "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc=\"",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "bucketId": "FtzysDykv0-9s9toWiZhdskAD67z",
      "title": "Bucket Task 1",
      "orderHint": "8585760017701920008P'",
      "assigneePriority": "",
      "percentComplete": 0,
      "startDateTime": null,
      "createdDateTime": "2021-07-06T20:59:35.4105517Z",
      "dueDateTime": null,
      "hasDescription": false,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "id": "KvamtRjaPkmPVy1rEA1r2skAOxcA",
      "createdBy": {
        "user": {
          "displayName": null,
          "id": "73829096-6f0a-4745-8f72-12a17bacadea"
        }
      },
      "appliedCategories": {},
      "assignments": {},
      "priority": 5
    },
    {
      "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc=\"",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "bucketId": "FtzysDykv0-9s9toWiZhdskAD67z",
      "title": "Bucket Task 2",
      "orderHint": "8585763504689506592PK",
      "assigneePriority": "8585763504089037251",
      "percentComplete": 0,
      "startDateTime": null,
      "createdDateTime": "2021-07-02T20:07:56.5738556Z",
      "dueDateTime": null,
      "hasDescription": false,
      "previewType": "automatic",
      "completedDateTime": null,
      "completedBy": null,
      "referenceCount": 0,
      "checklistItemCount": 0,
      "activeChecklistItemCount": 0,
      "conversationThreadId": null,
      "id": "BNWGt05mFUq1VV-cdK00aMkAH5nT",
      "createdBy": {
        "user": {
          "displayName": null,
          "id": "73829096-6f0a-4745-8f72-12a17bacadea"
        }
      },
      "appliedCategories": {},
      "assignments": {},
      "priority": 1
    }
  ];

  const taskListResponse: any = {
    "value": taskListResponseValue
  };

  const taskListBetaResponse: any = {
    "value": taskListResponseBetaValue
  };

  const bucketListResponseValue = [
    {
      "name": "Planner Bucket A",
      "id": "FtzysDykv0-9s9toWiZhdskAD67z"
    },
    {
      "name": "Planner Bucket 2",
      "id": "ZpcnnvS9ZES2pb91RPxQx8kAMLo5"
    }
  ];

  const bucketListResponse: any = {
    "value": bucketListResponseValue
  };

  const groupByDisplayNameResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": "0d0402ee-970f-4951-90b5-2f24519d2e40"
      }
    ]
  };

  const planResponse: any = {
    "value": [{
      "id": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "title": "My Planner Plan"
    }]
  };

  const plansInOwnerGroup: any = {
    "value": [
      {
        "title": "My Planner Plan",
        "id": "iVPMIgdku0uFlou-KLNg6MkAE1O2"
      },
      {
        "title": "Sample Plan",
        "id": "uO1bj3fdekKuMitpeJqaj8kADBxO"
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
  });

  beforeEach(() => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'&$select=id`) {
        return groupByDisplayNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/DjL5xiKO10qut8LQgztpKskABWna/plans?$select=id`) {
        return planResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans?$select=id,title`) {
        return plansInOwnerGroup;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets?$select=id,name`) {
        return bucketListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return taskListBetaResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return taskListBetaResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/me/planner/tasks`) {
        return taskListBetaResponse;
      }
      throw 'Invalid Request';
    });
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
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime']);
  });

  it('fails validation when both bucketId and bucketName are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketId: 'FtzysDykv0-9s9toWiZhdskAD67z',
        bucketName: 'Planner Bucket A'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucketName is specified without planId, planTitle, or rosterId', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketName: 'Planner Bucket A'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucketName is specified with planId, planTitle, and rosterId', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        planTitle: 'My Planner Plan',
        rosterId: 'DjL5xiKO10qut8LQgztpKskABWna'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when bucketName is specified with neither the planId, planTitle, nor rosterId', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        debug: true,
        bucketName: 'Planner Bucket A'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planId, planTitle, and rosterId are specified', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        planTitle: 'My Planner',
        rosterId: 'DjL5xiKO10qut8LQgztpKskABWna'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified without ownerGroupId or ownerGroupName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planTitle: 'My Planner Plan'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified with both ownerGroupId and ownerGroupName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planTitle: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when owner group id is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupId: 'Invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid planId is specified', async () => {
    const actual = await command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        bucketName: 'Planner Bucket A'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid planTitle and ownerGroupId are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid planTitle and ownerGroupName are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupName: 'My Planner Group'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName and planId are specified', async () => {
    const actual = await command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        bucketName: 'Planner Bucket A'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName, planTitle, and ownerGroupId are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when bucketName, planTitle, and ownerGroupName are specified', async () => {
    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupName: 'My Planner Group'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when no arguments are specified', async () => {
    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    const actual = await command.validate({
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when ownerGroupName not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return { value: [] };
      }
      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        planTitle: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }), new CommandError(`The specified group 'foo' does not exist.`));
  });

  it('fails validation when bucketName not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'&$select=id`) {
        return groupByDisplayNameResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/DjL5xiKO10qut8LQgztpKskABWna/plans?$select=id`) {
        return planResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans?$select=id,title`) {
        return plansInOwnerGroup;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks?$select=id,title`) {
        return taskListResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return taskListBetaResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return taskListBetaResponse;
      }
      if (opts.url === `https://graph.microsoft.com/beta/me/planner/tasks?$select=id,title`) {
        return taskListBetaResponse;
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets?$select=id,name`) {
        return { value: [] };
      }
      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        bucketName: 'foo',
        planTitle: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    }), new CommandError(`The specified bucket 'foo' does not exist.`));
  });

  it('lists planner tasks of the current logged in user', async () => {
    await command.action(logger, { options: {} });
    assert(loggerLogSpy.called);
  });

  it('correctly lists planner tasks with planTitle and ownerGroupId', async () => {
    const options: any = {
      planTitle: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with planTitle and ownerGroupName', async () => {
    const options: any = {
      planTitle: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with bucketId', async () => {
    const options: any = {
      bucketId: 'FtzysDykv0-9s9toWiZhdskAD67z'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with bucketName and planId', async () => {

    const options: any = {
      bucketName: 'Planner Bucket A',
      planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with bucketName, planTitle, and ownerGroupId', async () => {
    const options: any = {
      bucketName: 'Planner Bucket A',
      planTitle: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with bucketName, planTitle, and ownerGroupName', async () => {
    const options: any = {
      bucketName: 'Planner Bucket A',
      planTitle: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly lists planner tasks with bucketId and rosterId', async () => {
    const options: any = {
      bucketName: 'Planner Bucket A',
      rosterId: 'DjL5xiKO10qut8LQgztpKskABWna'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskListResponseBetaValue));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
