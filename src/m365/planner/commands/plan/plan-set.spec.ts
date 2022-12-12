import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./plan-set');

describe(commands.PLAN_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const id = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const title = 'Plan name';
  const ownerGroupName = 'Group name';
  const ownerGroupId = '00000000-0000-0000-0000-000000000002';
  const newTitle = 'New Title';
  const user = 'user@contoso.com';
  const userId = '00000000-0000-0000-0000-000000000000';
  const user1 = 'user1@contoso.com';
  const user1Id = '00000000-0000-0000-0000-000000000001';
  const shareWithUserNames = `${user},${user1}`;
  const shareWithUserIds = `${userId},${user1Id}`;
  const category21 = 'ToDo';
  const category25 = 'Urgent';

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

  const etagResponse = {
    "@odata.etag": "TestEtag"
  };

  const singleGroupsResponse = {
    value: [
      {
        id: ownerGroupId,
        displayName: ownerGroupName
      }
    ]
  };

  const singlePlansResponse = {
    value: [
      {
        '@odata.etag': 'abcdef',
        id: id,
        title: title,
        owner: ownerGroupId
      }
    ]
  };

  const planResponse = {
    "id": id,
    "title": title
  };

  const planDetailsResponse = {
    "sharedWith": {
      "00000000-0000-0000-0000-000000000000": true,
      "00000000-0000-0000-0000-000000000001": true
    },
    "categoryDescriptions": {
      "category1": null,
      "category2": null,
      "category3": null,
      "category4": null,
      "category5": null,
      "category6": null,
      "category7": null,
      "category8": null,
      "category9": null,
      "category10": null,
      "category11": null,
      "category12": null,
      "category13": null,
      "category14": null,
      "category15": null,
      "category16": null,
      "category17": null,
      "category18": null,
      "category19": null,
      "category20": null,
      "category21": category21,
      "category22": null,
      "category23": null,
      "category24": null,
      "category25": category25
    }
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      request.patch,
      accessToken.isAppOnlyAccessToken
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PLAN_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [{ options: ['id', 'title'] }]);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner']);
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: 'invalid guid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided when using title.', async () => {
    const actual = await command.validate({
      options: {
        title: title
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both ownerGroupId and ownerGroupName are specified when using title', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        ownerGroupName: ownerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if shareWithUserIds contains invalid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        shareWithUserIds: "invalid guid"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both shareWithUserIds and shareWithUserNames are specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        shareWithUserIds: shareWithUserIds,
        shareWithUserNames: shareWithUserNames
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when invalid category are specified', async () => {
    const actual = await command.validate({
      options: {
        id: id,
        category27: 'ToDo',
        category35: 'Urgent'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid title and ownerGroupId specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupName: ownerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, ownerGroupName, and shareWithUserIds specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        shareWithUserIds: shareWithUserIds
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when valid title, ownerGroupName, and validShareWithUserNames specified', async () => {
    const actual = await command.validate({
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        shareWithUserNames: shareWithUserNames
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when using app only access token', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        ownerGroupId: ownerGroupId
      }
    }), new CommandError('This command does not support application permissions.'));
  });

  it('correctly updates planner plan title with given id (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return planDetailsResponse;
      }

      return 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: id,
        newTitle: newTitle
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly updates planner plan shareWithUserNames with given title and ownerGroupName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(ownerGroupName)}'`) {
        return Promise.resolve(singleGroupsResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlansResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user1)}'&$select=id,userPrincipalName`) {
        return user1Response;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return planDetailsResponse;
      }

      return 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return outputResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        title: title,
        ownerGroupName: ownerGroupName,
        shareWithUserNames: shareWithUserNames
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly updates planner plan shareWithUserIds with given title and ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(ownerGroupName)}'`) {
        return Promise.resolve(singleGroupsResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlansResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return planDetailsResponse;
      }

      return 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return outputResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        title: title,
        ownerGroupId: ownerGroupId,
        shareWithUserIds: shareWithUserIds
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly updates planner plan categories with given id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return etagResponse;
      }

      return 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return outputResponse;
      }

      return 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: true,
        id: id,
        category21: category21,
        category25: category25
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('fails when an invalid user is specified', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(ownerGroupName)}'`) {
        return Promise.resolve(singleGroupsResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlansResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}`) {
        return planResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user)}'&$select=id,userPrincipalName`) {
        return userResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(user1)}'&$select=id,userPrincipalName`) {
        return { value: [] };
      }

      return 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${id}/details`) {
        return outputResponse;
      }

      return 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        title: title,
        ownerGroupName: ownerGroupName,
        shareWithUserNames: shareWithUserNames
      }
    }), new CommandError(`Cannot proceed with planner plan creation. The following users provided are invalid: ${user1}`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred.');
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred.'));
  });
});
