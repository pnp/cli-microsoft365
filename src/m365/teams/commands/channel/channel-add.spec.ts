import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './channel-add.js';
import { teams } from '../../../../utils/teams.js';

describe(commands.CHANNEL_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound,
      teams.getTeamIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are passed', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '00000000-0000-0000-0000-000000000000',
      teamName: 'Team Name',
      name: 'Architecture Discussion',
      description: 'Architecture'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if both channelId and channelName options are not passed', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'Architecture Discussion',
      description: 'Architecture'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: 'invalid GUID',
      name: 'Architecture Discussion',
      description: 'Architecture'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if unkown type is specified.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture Discussion',
      type: 'invalid'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if owner is not specified when creating private channel.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture Discussion',
      type: 'private'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if owner is specified when not creating private channel.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture Discussion',
      owner: 'John.Doe@contoso.com'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if owner is not specified when creating shared channel.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture Discussion',
      type: 'shared'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if owner is specified when not creating a private or shared channel.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture Discussion',
      owner: 'John.Doe@contoso.com'
    });
    assert.strictEqual(actual.success, false);
  });

  it('validates for a correct general channel input.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture',
      description: 'Architecture meeting'
    });
    assert.strictEqual(actual.success, true);
  });

  it('validates for a correct private channel input.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture',
      description: 'Architecture meeting',
      type: 'private',
      owner: 'john.doe@contoso.com'
    });
    assert.strictEqual(actual.success, true);
  });

  it('validates for a correct shared channel input.', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'Architecture',
      description: 'Architecture meeting',
      type: 'shared',
      owner: 'john.doe@contoso.com'
    });
    assert.strictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      name: 'test',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('fails to get team when team does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/me/joinedTeams') {
        return { value: [] };
      }

      throw 'The specified team does not exist in the Microsoft Teams';
    });

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion',
        description: 'Architecture'
      })
    } as any), new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('creates channel within the Microsoft Teams team in the tenant with description by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": "Architecture"
        };
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        debug: true,
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        description: 'Architecture'
      })
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": "Architecture"
    }));
  });

  it('creates channel within the Microsoft Teams team in the tenant without description by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion'
      })
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "description": null
    }));
  });

  it('creates private channel within the Microsoft Teams team by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "membershipType": "private"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'private',
        owner: 'john.doe@contoso.com'
      })
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "membershipType": "private"
    }));
  });

  it('creates shared channel within the Microsoft Teams team by team id', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "membershipType": "shared"
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion',
        type: 'shared',
        owner: 'john.doe@contoso.com'
      })
    });

    assert(loggerLogSpy.calledWith({
      "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
      "displayName": "Architecture Discussion",
      "membershipType": "shared"
    }));
  });

  it('creates channel within the Microsoft Teams team in the tenant by team name', async () => {
    sinon.stub(teams, 'getTeamIdByDisplayName').callsFake(async (name) => {
      if (name === 'Team Name') {
        return '00000000-0000-0000-0000-000000000000';
      }

      throw 'Invalid parameter passed to teams.getTeamIdByDisplayName: ' + name;
    });

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return {
          "id": "19:d9c63a6d6a2644af960d74ea927bdfb0@thread.skype",
          "displayName": "Architecture Discussion",
          "description": null
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        verbose: true,
        teamName: 'Team Name',
        name: 'Architecture Discussion'
      })
    });

    assert(postStub.calledOnce);
    assert.deepStrictEqual(postStub.firstCall.args[0].data, {
      membershipType: 'standard',
      displayName: 'Architecture Discussion'
    });
  });

  it('correctly handles error when adding a channel', async () => {
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'post').rejects(error);

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture Discussion'
      })
    }), new CommandError('An error has occurred'));
  });
});
