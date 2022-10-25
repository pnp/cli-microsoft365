import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { accessToken } from '../../../../utils/accessToken';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./plan-get');

describe(commands.PLAN_GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validTitle = 'Plan name';
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

  const planResponse = {
    "id": validId,
    "title": validTitle
  };

  const planDetailsResponse = {
    "sharedWith": {},
    "categoryDescriptions": {}
  };

  const outputResponse = {
    ...planResponse,
    ...planDetailsResponse
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PLAN_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner', '@odata.etag']);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('fails validation if the ownerGroupId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: invalidOwnerGroupId
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

  it('fails validation if neither the ownerGroupId nor ownerGroupName are provided with deprecated planTitle', async () => {
    const actual = await command.validate({
      options: {
        planTitle: validTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when id specified', async () => {
    const actual = await command.validate({
      options: {
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when title and valid ownerGroupId specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when title and valid ownerGroupName specified', async () => {
    const actual = await command.validate({
      options: {
        title: validTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [['id', 'title', 'planId', 'planTitle']]);
  });

  it('correctly get planner plan with given id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return Promise.resolve(planResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    await command.action(logger, {
      options: {
        debug: false,
        id: validId
      }
    });

    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with deprecated planId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}`) {
        return Promise.resolve(planResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    await command.action(logger, {
      options: {
        debug: false,
        planId: validId
      }
    });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given title and ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await command.action(logger, { options: options });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly get planner plan with given ownerGroupId and deprecated planTitle', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      planTitle: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await command.action(logger, { options: options });
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('fails validation when using app only access token', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        id: validId
      }
    }), new CommandError('This command does not support application permissions.'));
  });

  it('correctly get planner plan with given title and ownerGroupName', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf('/groups?$filter=displayName') > -1) {
        return Promise.resolve(singleGroupResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({
          "value": [
            planResponse
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validId}/details`) {
        return Promise.resolve(planDetailsResponse);
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupName: validOwnerGroupName
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(outputResponse));
  });

  it('correctly handles no plan found with given ownerGroupId', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve({ "value": [] });
      }

      return Promise.reject(`Invalid request ${opts.url}`);
    });

    const options: any = {
      debug: false,
      title: validTitle,
      ownerGroupId: validOwnerGroupId
    };

    await assert.rejects(command.action(logger, { options: options } as any));
    assert(loggerLogSpy.notCalled);
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred.');
    });


    await assert.rejects(command.action(logger, { options: { debug: false } }), new CommandError('An error has occurred.'));
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