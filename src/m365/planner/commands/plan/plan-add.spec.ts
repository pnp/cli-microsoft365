import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken, formatting, sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./plan-add');

describe(commands.PLAN_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const validId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTitle = 'Plan name';
  const validOwnerGroupName = 'Group name';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000002';
  const user = 'user@contoso.com';
  const userId = '00000000-0000-0000-0000-000000000000';
  const user1 = 'user1@contoso.com';
  const user1Id = '00000000-0000-0000-0000-000000000001';
  const validShareWithUserNames = `${user},${user1}`;
  const validShareWithUserIds = `${userId},${user1Id}`;

  const singleGroupResponse = {
    "value": [
      {
        "id": validOwnerGroupId,
        "displayName": validOwnerGroupName
      }
    ]
  };

  const planResponse = {
    "id": validId,
    "title": validTitle
  };

  const userResponse = {
    "value": [
      {
        "id": userId,
        "userPrincipalName": user
      }
    ]
  };

  const user1Response = {
    "value": [
      {
        "id": user1Id,
        "userPrincipalName": user1
      }
    ]
  };

  const planDetailsEtagResponse = {
    "@odata.etag": "TestEtag"
  };

  const planDetailsResponse = {
    "sharedWith": {
      "00000000-0000-0000-0000-000000000000": true,
      "00000000-0000-0000-0000-000000000001": true
    },
    "categoryDescriptions": {}
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
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
    assert.strictEqual(command.name.startsWith(commands.PLAN_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner']);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: 'no guid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if shareWithUserIds contains invalid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        shareWithUserIds: "no guid"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both shareWithUserIds and shareWithUserNames are specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        shareWithUserIds: validShareWithUserIds,
        shareWithUserNames: validShareWithUserNames
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid title and ownerGroupId specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, ownerGroupName, and shareWithUserIds specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        shareWithUserIds: validShareWithUserIds
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, ownerGroupName, and validShareWithUserNames specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId,
        shareWithUserNames: validShareWithUserNames
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when using app only access token', (done) => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    command.action(logger, {
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId
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

  it('correctly adds planner plan with given title with available ownerGroupId', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve(planResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(planResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly adds planner plan with given title with available ownerGroupName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve(singleGroupResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve(planResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupName: validOwnerGroupName
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(planResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly adds planner plan with given title with ownerGroupId and shareWithUserIds', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsEtagResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve(planResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId,
      shareWithUserIds: validShareWithUserIds
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly adds planner plan with given title with ownerGroupId and shareWithUserNames', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsEtagResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`) {
        return Promise.resolve(userResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user1)}'&$select=id,userPrincipalName`) {
        return Promise.resolve(user1Response);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve(planResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId,
      shareWithUserNames: validShareWithUserNames
    };

    command.action(logger, { options: options } as any, () => {
      try {
        assert(loggerLogSpy.calledWith(outputResponse));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when an invalid user is specified', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsEtagResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`) {
        return Promise.resolve(userResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user1)}'&$select=id,userPrincipalName`) {
        return Promise.resolve({ value: [] });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans`) {
        return Promise.resolve(planResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId,
      shareWithUserNames: validShareWithUserNames
    };

    command.action(logger, { options: options } as any, (err: any) => {
      try {
        assert.strictEqual(err.message, `Cannot proceed with planner plan creation. The following users provided are invalid : ${user1}`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles API OData error', (done) => {
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