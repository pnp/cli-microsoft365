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
const command: Command = require('./task-add');

describe(commands.TASK_ADD, () => {
  const taskAddResponse = {
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
    "appliedCategories": {},
    "assignments": {}
  };

  const taskAddResponseWithDetails = {
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622710787367671",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:03:26.7408136Z",
    "dueDateTime": null,
    "hasDescription": true,
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
    "appliedCategories": {},
    "assignments": {},
    "description": "My Task Description",
    "references": {},
    "checklist": {}
  };

  const taskAddResponseWithAssignments = {
    "@odata.etag": "W/\"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBARCc=\"",
    "planId": "8QZEH7b3wkS_bGQobscsM5gADCBb",
    "bucketId": "IK8tuFTwQEa5vTonM7ZMRZgAKdno",
    "title": "My Planner Task",
    "orderHint": "8585622689173829649",
    "assigneePriority": "",
    "percentComplete": 0,
    "startDateTime": null,
    "createdDateTime": "2021-12-12T19:39:28.0946158Z",
    "dueDateTime": null,
    "hasDescription": false,
    "previewType": "automatic",
    "completedDateTime": null,
    "completedBy": null,
    "referenceCount": 0,
    "checklistItemCount": 0,
    "activeChecklistItemCount": 0,
    "conversationThreadId": null,
    "id": "mEsX2erws0CHP4PUn_ZlNJgAI2VQ",
    "createdBy": {
      "user": {
        "displayName": null,
        "id": "dd8b99a7-77c6-4238-a609-396d27844921"
      }
    },
    "appliedCategories": {},
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
        "id": "0d0402ee-970f-4951-90b5-2f24519d2e40",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2021-06-08T11:04:45Z",
        "creationOptions": [],
        "description": "My Planner Group",
        "displayName": "My Planner Group",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "isAssignableToRole": null,
        "mail": "MyPlannerGroup@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "My Planner Group",
        "membershipRule": null,
        "membershipRuleProcessingState": null,
        "onPremisesDomainName": null,
        "onPremisesLastSyncDateTime": null,
        "onPremisesNetBiosName": null,
        "onPremisesSamAccountName": null,
        "onPremisesSecurityIdentifier": null,
        "onPremisesSyncEnabled": null,
        "preferredDataLocation": null,
        "preferredLanguage": null,
        "proxyAddresses": [
          "SPO:SPO_e13f6193-fb01-43e8-8e8d-557796b82ebf@SPO_cc6fafe9-dd93-497c-b521-1d971b1471c7",
          "SMTP:MyPlannerGroup@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2021-06-08T11:04:45Z",
        "resourceBehaviorOptions": [],
        "resourceProvisioningOptions": [],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-218366702-1230083855-573552016-1076796785",
        "theme": null,
        "visibility": "Private",
        "onPremisesProvisioningErrors": []
      }
    ]
  };

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

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
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return taskAddResponse;
      }
      throw 'Invalid request';
    });

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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if neither the planId nor planTitle are provided.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both planId and planTitle are specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        planTitle: 'My Planner',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified without ownerGroupId or ownerGroupName', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planTitle: 'My Planner Plan',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when planTitle is specified with both ownerGroupId and ownerGroupName', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planTitle: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid title, planId, and bucketId specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, planTitle, and ownerGroupId are specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planTitle: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, planTitle, ownerGroupName, and bucketId are specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planTitle: 'My Planner Plan',
        ownerGroupName: 'My Planner Group',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planTitle: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the bucketId nor bucketName are provided.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both bucketId and bucketName are specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        bucketName: 'My Bucket'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if startDateTime contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        startDateTime: '2021-99-99'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if dueDateTime contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        dueDateTime: '2021-99-99'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if percentComplete contains invalid format.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        percentComplete: 'Not A Number'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if percentComplete is not between 0 and 100.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        percentComplete: 599
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if assignedToUserIds contains invalid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-test"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both assignedToUserIds and assignedToUserNames are specified', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        assignedToUserIds: "2e42fe76-3f42-4884-b325-aefd7a905446,8d1ff29c-a6f4-4786-b316-eb6030e1a09e",
        assignedToUserNames: "Allan.Carroll@contoso.onmicrosoft.com,Ida.Stevens@contoso.onmicrosoft.com"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if incorrect appliedCategory is specified.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        appliedCategories: "category1,category9"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority lower than 0 is specified.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        priority: -1
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if incorrect previewType is specified.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        previewType: "test"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority higher than 10 is specified.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        priority: 11
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if priority is specified which is a number with decimals.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        priority: 5.6
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if unknown priority label is specified.', async () => {
    const actual = await command.validate({
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        priority: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('correctly adds planner task with title, planId, and bucketId', async () => {
    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponse));
  });

  it('correctly adds planner bucket with title, bucketId, planTitle, and ownerGroupName', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans`) {
        return {
          value: [
            {
              "owner": "0d0402ee-970f-4951-90b5-2f24519d2e40",
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('My Planner Group')}'`) {
        return groupByDisplayNameResponse;
      }

      throw 'Invalid request';
    });

    const options: any = {
      title: 'My Planner Task',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      planTitle: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponse));
  });

  it('correctly adds planner task with title, bucketId, planTitle, and ownerGroupId', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/0d0402ee-970f-4951-90b5-2f24519d2e40/planner/plans`) {
        return {
          value: [
            {
              "createdBy": {
                "application": {
                  "id": "95e27074-6c4a-447a-aa24-9d718a0b86fa"
                },
                "user": {
                  "id": "ebf3b108-5234-4e22-b93d-656d7dae5874"
                }
              },
              "createdDateTime": "2015-03-30T18:36:49.2407981Z",
              "owner": "ebf3b108-5234-4e22-b93d-656d7dae5874",
              "title": "My Planner Plan",
              "id": "8QZEH7b3wkS_bGQobscsM5gADCBb"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const options: any = {
      title: 'My Planner Task',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      planTitle: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponse));
  });

  it('correctly adds planner task with title, planId, and bucketName', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/8QZEH7b3wkS_bGQobscsM5gADCBb/buckets`) {
        return {
          value: [
            {
              "name": "My Planner Bucket",
              "planId": "2txjA-BMZEq-bKi6Wfj5aGQAB1OJ",
              "orderHint": "85752723360752+",
              "id": "IK8tuFTwQEa5vTonM7ZMRZgAKdno"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketName: 'My Planner Bucket'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponse));
  });

  it('correctly adds planner task with title, bucketId, planId, and assignedToUserIds', async () => {
    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      assignedToUserIds: '949b16c1-a032-453e-a8ae-89a52bfc1d8a'
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponse));
  });

  it('correctly adds planner task with title, bucketId, planId, assignedToUserNames, and appliedCategories', async () => {
    sinonUtil.restore(request.get);
    sinonUtil.restore(request.post);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      assignedToUserNames: 'user@contoso.onmicrosoft.com',
      appliedCategories: "category1,category3"
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter('user@contoso.onmicrosoft.com')}'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return taskAddResponseWithAssignments;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponseWithAssignments));
  });

  it('correctly adds planner task with title, bucketId, planId, assignedToUserNames, and appliedCategories split with space', async () => {
    sinonUtil.restore(request.get);
    sinonUtil.restore(request.post);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      assignedToUserNames: 'user@contoso.onmicrosoft.com',
      appliedCategories: "category1 category2"
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter('user@contoso.onmicrosoft.com')}'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return taskAddResponseWithAssignments;
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponseWithAssignments));
  });

  it('correctly adds planner task with title, bucketId, planId, and description', async () => {
    sinonUtil.restore(request.get);
    sinonUtil.restore(request.patch);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      description: 'My Task Description'
    };

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/Z-RLQGfppU6H3663DBzfs5gAMD3o/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return {
          "@odata.etag": "TestEtag"
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks`) {
        return taskAddResponseWithDetails;
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/Z-RLQGfppU6H3663DBzfs5gAMD3o/details`) {
        return {
          "description": "My Task Description",
          "references": {},
          "checklist": {}
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(taskAddResponseWithDetails));
  });

  it('uses correct value for urgent priority', async () => {
    sinonUtil.restore(request.post);
    const requestPostStub = sinon.stub(request, 'post');
    requestPostStub.resolves(taskAddResponseWithAssignments);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      priority: 'Urgent'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPostStub.lastCall.args[0].data.priority, 1);
  });

  it('uses correct value for important priority', async () => {
    sinonUtil.restore(request.post);
    const requestPostStub = sinon.stub(request, 'post');
    requestPostStub.resolves(taskAddResponseWithAssignments);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      priority: 'Important'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPostStub.lastCall.args[0].data.priority, 3);
  });

  it('uses correct value for medium priority', async () => {
    sinonUtil.restore(request.post);
    const requestPostStub = sinon.stub(request, 'post');
    requestPostStub.resolves(taskAddResponseWithAssignments);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      priority: 'Medium'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPostStub.lastCall.args[0].data.priority, 5);
  });

  it('uses correct value for low priority', async () => {
    sinonUtil.restore(request.post);
    const requestPostStub = sinon.stub(request, 'post');
    requestPostStub.resolves(taskAddResponseWithAssignments);

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      priority: 'Low'
    };

    await command.action(logger, { options: options } as any);
    assert.strictEqual(requestPostStub.lastCall.args[0].data.priority, 9);
  });

  it('fails when no bucket is found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/8QZEH7b3wkS_bGQobscsM5gADCBb/buckets`) {
        return {
          value: []
        };
      }

      throw 'Invalid request';
    });

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketName: 'My Planner Bucket'
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('The specified bucket does not exist'));
  });

  it('fails when an invalid user is specified', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return {
          value: [
            {
              id: '949b16c1-a032-453e-a8ae-89a52bfc1d8a',
              userPrincipalName: 'user@contoso.onmicrosoft.com'
            }
          ]
        };
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq 'user2%40contoso.onmicrosoft.com'&$select=id,userPrincipalName`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    const options: any = {
      title: 'My Planner Task',
      planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
      bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
      assignedToUserNames: 'user@contoso.onmicrosoft.com,user2@contoso.onmicrosoft.com'
    };

    await assert.rejects(command.action(logger, { options: options } as any), new CommandError('Cannot proceed with planner task creation. The following users provided are invalid : user2@contoso.onmicrosoft.com'));
  });

  it('fails validation when ownerGroupName not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return { value: [] };
      }
      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: 'My Planner Bucket',
        planTitle: 'My Planner Plan',
        ownerGroupName: 'foo'
      }
    }), new CommandError(`The specified group 'foo' does not exist.`));
  });

  it('fails validation when task details endpoint fails', async () => {
    sinonUtil.restore(request.get);

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/Z-RLQGfppU6H3663DBzfs5gAMD3o/details`) {
        throw 'Error fetching task details';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: 'My Planner Task',
        planId: '8QZEH7b3wkS_bGQobscsM5gADCBb',
        bucketId: 'IK8tuFTwQEa5vTonM7ZMRZgAKdno',
        description: 'My Task Description'
      }
    }), new CommandError(`Error fetching task details`));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});