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
const command: Command = require('./plan-remove');

describe(commands.PLAN_REMOVE, () => {
  const validPlanTitle = 'My Plan';
  const validPlanId = 'opb7bchfZUiFbVWEPL7jPGUABW7f';
  const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';
  const validOwnerGroupName = 'HR';

  const singlePlanResponse = {
    '@odata.etag': 'abcdef',
    title: validPlanTitle,
    owner: validOwnerGroupId,
    id: validPlanId
  };

  const singleGroupsResponse = {
    value: [
      {
        id: validOwnerGroupId,
        displayName: validOwnerGroupName
      }
    ]
  };

  const singlePlansResponse = {
    value: [
      {
        '@odata.etag': 'abcdef',
        id: validPlanId,
        title: validPlanTitle,
        owner: validOwnerGroupId
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

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
    promptOptions = undefined;
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(Cli, 'prompt').callsFake(async (options: any) => {
      promptOptions = options;
      return { continue: false };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.delete,
      accessToken.isAppOnlyAccessToken,
      Cli.prompt
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
    assert.strictEqual(command.name, commands.PLAN_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation when id and ownerGroupId is specified', async () => {
    const actual = await command.validate({
      options: {
        id: validPlanId,
        ownerGroupId: validOwnerGroupId
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified with both ownerGroupName and ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupId: validOwnerGroupId,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified without ownerGroupName or ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when title is specified with invalid ownerGroupId', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with id', async () => {
    const actual = await command.validate({
      options: {
        id: validPlanId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct input with title', async () => {
    const actual = await command.validate({
      options: {
        title: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the specified plan when confirm option not passed with id', async () => {
    await command.action(logger, {
      options: {
        id: validPlanId
      }
    });

    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }

    assert(promptIssued);
  });

  it('aborts removing the specified plan when confirm option not passed and prompt not confirmed', async () => {
    const deleteSpy = sinon.spy(request, 'delete');
    await command.action(logger, {
      options: {
        id: validPlanId
      }
    });
    assert(deleteSpy.notCalled);
  });

  it('fails validation when using app only access token', async () => {
    sinonUtil.restore(accessToken.isAppOnlyAccessToken);
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {
        id: validPlanId
      }
    }), new CommandError('This command does not support application permissions.'));
  });

  it('Correctly deletes plan by id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return Promise.resolve(singlePlanResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });

    await command.action(logger, {
      options: {
        id: validPlanId,
        confirm: true
      }
    });
  });

  it('Correctly deletes plan by title', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validOwnerGroupName)}'`) {
        return Promise.resolve(singleGroupsResponse);
      }
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlansResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: validPlanTitle,
        ownerGroupName: validOwnerGroupName
      }
    });
  });

  it('Correctly deletes plan by title with group id', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return Promise.resolve(singlePlansResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid Request');
    });
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: true }
    ));

    await command.action(logger, {
      options: {
        title: validPlanTitle,
        ownerGroupId: validOwnerGroupId
      }
    });
  });

  it('correctly handles random API error', async () => {
    sinon.stub(request, 'get').callsFake(() => Promise.resolve(singlePlanResponse));
    sinon.stub(request, 'delete').callsFake(() => Promise.reject('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: {
        id: validPlanId,
        confirm: true
      }
    }), new CommandError("An error has occurred"));
  });
});
