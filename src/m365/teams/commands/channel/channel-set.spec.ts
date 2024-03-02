import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './channel-set.js';

describe(commands.CHANNEL_SET, () => {
  const id = '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype';
  const name = 'channelName';
  const teamId = 'd66b8110-fcad-49e8-8159-0d488ddb7656';
  const teamName = 'Team Name';
  const newName = 'New Review';
  const description = 'This is a new description';

  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_SET);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly validates the arguments', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        name: name,
        newName: newName,
        description: description
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid',
        name: name,
        newName: newName,
        description: description
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not valid', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        id: 'invalid',
        newName: newName,
        description: description
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when name is General', async () => {
    const actual = await command.validate({
      options: {
        teamId: teamId,
        name: 'General',
        newName: newName,
        description: description
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails to patch channel when channel does not exists', async () => {
    const errorMessage = 'The specified channel does not exist in this Microsoft Teams team';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`) {
        return { value: [] };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        debug: true,
        teamId: teamId,
        name: name,
        newName: newName,
        description: description
      }
    }), new CommandError(errorMessage));
  });

  it('correctly patches channel updates by teamId and name', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'`) {
        return {
          value:
            [
              {
                "id": id,
                "displayName": "Review",
                "description": "Updated by CLI"
              }]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: newName, description: description })
      ) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamId: teamId,
        name: name,
        newName: newName,
        description: description
      }
    });
  });

  it('correctly patches channel updates by teamName and id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'`) {
        return {
          value: [
            {
              "id": teamId,
              "displayName": teamName,
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if ((opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter(teamId)}/channels/${formatting.encodeQueryParameter(id)}`) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: newName, description: description })
      ) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamName: teamName,
        id: id,
        newName: newName,
        description: description
      }
    });
  });

  it('fails when team name does not exist', async () => {
    const errorMessage = 'The specified team does not exist';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": []
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        id: id,
        teamName: teamName,
        newName: newName,
        description: description,
        force: true
      }
    }), new CommandError(errorMessage));
  });
});
