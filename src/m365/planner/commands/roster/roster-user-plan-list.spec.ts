import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Cli } from '../../../../cli/Cli';
import { accessToken } from '../../../../utils/accessToken';
const command: Command = require('./roster-user-plan-list');

describe(commands.ROSTER_USER_PLAN_LIST, () => {
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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      accessToken.isAppOnlyAccessToken,
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      session.getId
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.ROSTER_USER_PLAN_LIST);
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

    await command.action(logger, { options: { debug: true } });
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

    await command.action(logger, { options: { debug: true, userName: userName } });
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

    await command.action(logger, { options: { debug: true, userId: userId } });
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
    await assert.rejects(command.action(logger, { options: { userId: userId } } as any), new CommandError(`The options 'userId' or 'userName' cannot be used when obtaining Microsoft Planner Roster plans using delegated permissions`));
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