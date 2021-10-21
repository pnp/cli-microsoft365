import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./task-list');

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

  const taskListResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerTask)",
    "@odata.count": 2,
    "value": taskListResponseValue
  };

  const bucketListResponseValue = [
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket A",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585768731950308408",
      "id": "FtzysDykv0-9s9toWiZhdskAD67z"
    },
    {
      "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
      "name": "Planner Bucket 2",
      "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
      "orderHint": "8585784565[8",
      "id": "ZpcnnvS9ZES2pb91RPxQx8kAMLo5"
    }
  ];

  const bucketListResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerBucket)",
    "@odata.count": 2,
    "value": bucketListResponseValue
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

  const plansInOwnerGroup: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#planner/plans",
    "@odata.count": 2,
    "value": [
      {
        "@odata.etag": "W/\"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBASCc=\"",
        "createdDateTime": "2021-06-08T12:24:57.3312829Z",
        "owner": "f3f985d0-a4e0-4891-83f6-08d88bf44e5e",
        "title": "My Planner Plan",
        "id": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "73829066-5f0a-4745-8f72-12a17bacadea"
          },
          "application": {
            "displayName": null,
            "id": "09abbdfd-ed25-47ee-a2d9-a627aa1c90f3"
          }
        }
      },
      {
        "@odata.etag": "W/\"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBASCc=\"",
        "createdDateTime": "2021-06-08T12:25:09.3751058Z",
        "owner": "f3f985d0-a4e0-4891-83f6-08d88bf44e5e",
        "title": "Sample Plan",
        "id": "uO1bj3fdekKuMitpeJqaj8kADBxO",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "73829066-5f0a-4745-8f72-12a17bacadea"
          },
          "application": {
            "displayName": null,
            "id": "09abbdfd-ed25-47ee-a2d9-a627aa1c90f3"
          }
        }
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
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) {
        return Promise.resolve(bucketListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      return Promise.reject('Invalid Request');
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
  });

  afterEach(() => {
    Utils.restore([
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.restoreAuth
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'startDateTime', 'dueDateTime', 'completedDateTime']);
  });

  it('fails validation when both bucketId and bucketName are specified', (done) => {
    const actual = command.validate({
      options: {
        bucketId: 'FtzysDykv0-9s9toWiZhdskAD67z',
        bucketName: 'Planner Bucket A'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when bucketName is specified without planId or planName', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when bucketName is specified with both planId and planName', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        planName: 'My Planner Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when bucketName is specified with neither the planId nor planName', (done) => {
    const actual = command.validate({
      options: {
        debug: true,
        bucketName: 'Planner Bucket A'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both planId and planName are specified', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        planName: 'My Planner'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified without ownerGroupId or ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planName: 'My Planner Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified with both ownerGroupId and ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        bucketName: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid planId is specified', (done) => {
    const actual = command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid planName and ownerGroupId are specified', (done) => {
    const actual = command.validate({
      options: {
        planName: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid planName and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        planName: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when bucketName and planId are specified', (done) => {
    const actual = command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        bucketName: 'Planner Bucket A'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when bucketName, planName, and ownerGroupId are specified', (done) => {
    const actual = command.validate({
      options: {
        planName: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when bucketName, planName, and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        planName: 'My Planner Plan',
        bucketName: 'Planner Bucket A',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when no arguments are specified', (done) => {
    const actual = command.validate({
      options: {
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        bucketName: 'Planner Bucket A'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        planName: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when ownerGroupName not found', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
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

  it('fails validation when bucketName not found', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        bucketName: 'foo',
        planName: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified bucket does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when planName not found', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) {
        return Promise.resolve(bucketListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/me/planner/tasks`) {
        return Promise.resolve(taskListResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')`) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
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

  it('lists planner tasks of the current logged in user', (done) => {
    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(loggerLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with planName and ownerGroupId', (done) => {
    const options: any = {
      debug: false,
      planName: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with planName and ownerGroupName', (done) => {
    const options: any = {
      debug: false,
      planName: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with bucketId', (done) => {
    const options: any = {
      debug: false,
      bucketId: 'FtzysDykv0-9s9toWiZhdskAD67z'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with bucketName and planId', (done) => {
    const options: any = {
      debug: false,
      bucketName: 'Planner Bucket A',
      planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with bucketName, planName, and ownerGroupId', (done) => {
    const options: any = {
      debug: false,
      bucketName: 'Planner Bucket A',
      planName: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly lists planner tasks with bucketName, planName, and ownerGroupName', (done) => {
    const options: any = {
      debug: false,
      bucketName: 'Planner Bucket A',
      planName: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(taskListResponseValue));
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