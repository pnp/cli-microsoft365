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

  const multiplePlanResponse = {
    "value": [
      {
        "id": validPlanId,
        "title": validPlanName
      },
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

  const singleBucketByIdResponse = {
    "@odata.etag": "W/\"JzEtQnVja2V0QEBAQEBAQEBAQEBAQEBARCc=\"",
    "name": validBucketName,
    "id": validBucketId
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

  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.BUCKET_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when no option is specified', () => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and name are specified', () => {
    const actual = command.validate({
      options: {
        id: validBucketId,
        name: validBucketName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when id and plan details are specified', () => {
    const actual = command.validate({
      options: {
        id: validBucketId,
        planId: validPlanId
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used without plan id or planname', () => {
    const actual = command.validate({
      options: {
        name: validBucketName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used with both plan id and planname', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        planName: validPlanName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan name is used without owner group name or owner group id', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        planName: validPlanName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used with both owner group name and owner group id', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
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
        name: validBucketName,
        planName: validPlanName,
        ownerGroupId: invalidOwnerGroupId
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group name', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group id', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', () => {
    const actual = command.validate({
      options: {
        id: validBucketId
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', () => {
    const actual = command.validate({
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation when no groups found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve({"value": []});
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified owner group ${validOwnerGroupName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple groups found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve(multipleGroupResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple owner groups with name ${validOwnerGroupName} found: ${multipleGroupResponse.value.map(x => x.id)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no plans found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=owner eq '${validOwnerGroupId}'`) {
        return Promise.resolve({"value": []});
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupId: validOwnerGroupId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified plan ${validPlanName} does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when multiple plans found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=owner eq '${validOwnerGroupId}'`) {
        return Promise.resolve(multiplePlanResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupId: validOwnerGroupId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple plans with name ${validPlanName} found: ${multiplePlanResponse.value.map(x => x.id)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation when no buckets found', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve({"value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
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
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(multipleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        name: validBucketName,
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

  it('Correctly gets bucket by id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve(singleBucketByIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        id: validBucketId
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Correctly gets bucket by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve(singleGroupResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans?$filter=owner eq '${validOwnerGroupId}'`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(singleBucketByNameResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve(singleBucketByIdResponse);
      }

      return Promise.reject('Invalid Request');
    });
    
    command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanName,
        ownerGroupName: validOwnerGroupName
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined', err?.message);
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