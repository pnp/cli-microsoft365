import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./bucket-get');

describe(commands.BUCKET_GET, () => {
  const bucketGetResponseValue = {
    "name": "Planner Bucket A",
    "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
    "orderHint": "85752723360752+",
    "id": "FtzysDykv0-9s9toWiZhdskAD67z"
  };
  
  const bucketsInPlan = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.plannerBucket)",
    "@odata.count": 2,
    "value":[
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBAcCc=\"",
        "name": "Planner Bucket A",
        "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
        "orderHint": "858553C",
        "id": "FtzysDykv0-9s9toWiZhdskAD67z"
      },
      {
        "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBAcCc=\"",
        "name": "Planner Bucket B",
        "planId": "iVPMIgdku0uFlou-KLNg6MkAE1O2",
        "orderHint": "8585567758821058646P:",
        "id": "4Hb4fpEt2kiBIIKXPWPSs5gAL-RL"
      }
    ]
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
        return Promise.resolve(bucketsInPlan);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/FtzysDykv0-9s9toWiZhdskAD67z`) {
        return Promise.resolve(bucketGetResponseValue);
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
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.BUCKET_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'name', 'planId', 'orderHint']);
  });

  it('fails validation if neither the id nor title are provided.', (done) => {
    const actual = command.validate({
      options: {
        debug: true
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both Id and title are specified', (done) => {
    const actual = command.validate({
      options: {
        id: 'FtzysDykv0-9s9toWiZhdskAD67z',
        title: 'Planner Bucket A'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified without ownerGroupId or ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planName: 'My Planner Plan'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName is specified with both ownerGroupId and ownerGroupName', (done) => {
    const actual = command.validate({
      options: {
        title: "Planner Bucket A",
        planName: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid planId with title is specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid title, planName, ownerGroupId are specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid title, planName and ownerGroupName are specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupName: 'My Planner Group'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails validation if the ownerGroupId is not a valid guid.', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planName: 'My Planner Plan',
        ownerGroupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('correctly gets planner buckets with Id', (done) => {
    const options: any = {
      debug: false,
      id: 'FtzysDykv0-9s9toWiZhdskAD67z'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(bucketGetResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly gets planner bucket with title, planName and ownerGroupName', (done) => {
    const options: any = {
      debug: false,
      title: 'Planner Bucket A',
      planName: 'My Planner Plan',
      ownerGroupName: 'My Planner Group'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(bucketGetResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly gets planner bucket with title, planId', (done) => {
    const options: any = {
      debug: false,
      title: 'Planner Bucket A',
      planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(bucketGetResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('correctly gets planner bucket with title, planName, ownerGroupId ', (done) => {
    const options: any = {
      debug: false,
      title: 'Planner Bucket A',
      planName: 'My Planner Plan',
      ownerGroupId: '0d0402ee-970f-4951-90b5-2f24519d2e40'
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(bucketGetResponseValue));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when ownerGroupName not found', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'Planner Bucket A',
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

  it('fails validation if neither the planId nor planName are provided.', (done) => {
    const actual = command.validate({
      options: {
        debug: true,
        title: 'Planner Bucket A'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when both planId and planName are specified', (done) => {
    const actual = command.validate({
      options: {
        title: 'Planner Bucket A',
        planId: 'iVPMIgdku0uFlou-KLNg6MkAE1O2',
        planName: 'My Planner'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation when planName not found', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')`) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'Planner Bucket A',
        planName: 'My Planner',
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

  it('fails validation when bucket is not found', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('My Planner Group')}'`) {
        return Promise.resolve(groupByDisplayNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=(owner eq '${encodeURIComponent('0d0402ee-970f-4951-90b5-2f24519d2e40')}')`) {
        return Promise.resolve(plansInOwnerGroup);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/iVPMIgdku0uFlou-KLNg6MkAE1O2/buckets`) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: false,
        title: 'foo',
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

  it('correctly handles API OData error', (done) => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject("An error has occurred.");
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
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