import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { accessToken } from '../../../../utils/accessToken.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './roster-plan-list.js';

describe(commands.ROSTER_PLAN_LIST, () => {
  const userId = '59f80e08-24b1-41f8-8586-16765fd830d3';
  const userName = 'john.doe@contoso.com';
  const rosterUserPlanListResponse = {
    "value": [
      {
        "createdDateTime": "2023-04-06T14:41:49.8676617Z",
        "owner": "59f80e08-24b1-41f8-8586-16765fd830d3",
        "title": "My Planner Plan",
        "creationSource": null,
        "id": "_5GY9MJpZU2vb3DC46CP3MkACr8m",
        "createdBy": {
          "user": {
            "displayName": null,
            "id": "59f80e08-24b1-41f8-8586-16765fd830d3"
          },
          "application": {
            "displayName": null,
            "id": "31359c7f-bd7e-475c-86db-fdb8c937548e"
          }
        },
        "container": {
          "containerId": "_5GY9MJpZU2vb3DC46CP3MkACr8m",
          "type": "unknownFutureValue",
          "url": "https://graph.microsoft.com/beta/planner/rosters/_5GY9MJpZU2vb3DC46CP3MkACr8m"
        },
        "contexts": {},
        "sharedWithContainers": []
      }
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.active = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
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
      accessToken.isAppOnlyAccessToken,
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.active = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_PLAN_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if userId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { userId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userId defined', async () => {
    const actual = await command.validate({ options: { userId: userId } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if userName is not a valid UPN', async () => {
    const actual = await command.validate({ options: { userName: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input with a userName defined', async () => {
    const actual = await command.validate({ options: { userName: userName } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'createdDateTime', 'owner']);
  });

  it('retrieves all planner plans contained in roster where current logged in user is member of', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/planner/rosterPlans`) {
        return rosterUserPlanListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true } });
    assert(loggerLogSpy.calledWith(rosterUserPlanListResponse.value));
  });

  it('retrieves all planner plans contained in roster where specific user is member of by its UPN', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userName}/planner/rosterPlans`) {
        return rosterUserPlanListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userName: userName } });
    assert(loggerLogSpy.calledWith(rosterUserPlanListResponse.value));
  });

  it('retrieves all planner plans contained in roster where specific user is member of by its Id', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/${userId}/planner/rosterPlans`) {
        return rosterUserPlanListResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { verbose: true, userId: userId } });
    assert(loggerLogSpy.calledWith(rosterUserPlanListResponse.value));
  });

  it('throws an error when using application permissions and no option is specified', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(true);

    await assert.rejects(command.action(logger, {
      options: {}
    }), new CommandError(`Specify at least 'userId' or 'userName' when using application permissions.`));
  });

  it('throws an error when passing userId using delegated permissions', async () => {
    sinon.stub(accessToken, 'isAppOnlyAccessToken').returns(false);
    await assert.rejects(command.action(logger, {
      options: { userId: userId }
    }), new CommandError(`The options 'userId' or 'userName' cannot be used when obtaining Microsoft Planner Roster plans using delegated permissions.`));
  });

  it('handles error when retrieving all planner plans contained in roster', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/planner/rosterPlans`) {
        throw { error: { error: { message: 'An error has occurred' } } };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} }), new CommandError('An error has occurred'));
  });
});