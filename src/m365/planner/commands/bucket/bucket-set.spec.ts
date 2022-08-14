import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./bucket-set');

describe(commands.BUCKET_SET, () => {
  const validBucketId = 'vncYUXCRBke28qMLB-d4xJcACtNz';
  const validBucketName = 'Bucket name';
  const validOrderHint = '8585513699476931356P;';
  const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
  const validPlanTitle = 'Plan title';
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
  let commandInfo: CommandInfo;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.BUCKET_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'name']]);
  });

  it('fails validation id when id and plan details are specified', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId,
        planId: validPlanId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used without plan id or planTitle', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used with both plan id and planTitle', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        planTitle: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan name is used without owner group name or owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is used with both owner group name and owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
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
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: invalidOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when plan id is used with owner group id', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planId: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when new name or order hint is not specified', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validBucketId,
        newName: 'New name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with name', async () => {
    const actual = await command.validate({
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        newName: 'New name'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when using app only access token', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        id: validBucketId,
        planId: validPlanId,
        newName: validBucketName
      }
    }), new CommandError('This command does not support application permissions.'));
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
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
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
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }), new CommandError(`Multiple groups with name '${validOwnerGroupName}' found: ${multipleGroupResponse.value.map(x => x.id)}.`));
  });

  it('fails validation when no buckets found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`The specified bucket ${validBucketName} does not exist`));
  });

  it('fails validation when multiple buckets found', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(multipleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });

    await assert.rejects(command.action(logger, {
      options: {
        name: validBucketName,
        planId: validPlanId
      }
    }), new CommandError(`Multiple buckets with name ${validBucketName} found: ${multipleBucketByNameResponse.value.map(x => x.id)}`));
  });

  it('Correctly updates bucket by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve(singleBucketByIdResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        id: validBucketId,
        newName: validBucketName
      }
    }));
  });

  it('Correctly updates bucket by name', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(validOwnerGroupName)}'`) {
        return Promise.resolve(singleGroupResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(singleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupName: validOwnerGroupName,
        newName: 'New bucket name',
        orderHint: validOrderHint
      }
    }));
  });

  it('Correctly updates bucket by name with group ID', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(singleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planTitle: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        newName: 'New bucket name',
        orderHint: validOrderHint
      }
    }));
  });

  it('Correctly updates bucket by name with deprecated planName', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlanResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}/buckets`) {
        return Promise.resolve(singleBucketByNameResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/buckets/${validBucketId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await assert.doesNotReject(command.action(logger, {
      options: {
        name: validBucketName,
        planName: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        newName: 'New bucket name',
        orderHint: validOrderHint,
        verbose: true
      }
    }));
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