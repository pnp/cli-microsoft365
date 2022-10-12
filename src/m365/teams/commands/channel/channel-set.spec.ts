import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./channel-set');

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
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_SET), true);
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

  it('defines correct option sets', () => {
    const optionSets = command.optionSets;
    assert.deepStrictEqual(optionSets, [
      ['id', 'name'],
      ['teamId', 'teamName']
    ]);
  });

  it('fails to patch channel when channel does not exists', async () => {
    const errorMessage = 'The specified channel does not exist in the Microsoft Teams team';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels?$filter=displayName eq '${encodeURIComponent(name)}'`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels?$filter=displayName eq '${encodeURIComponent(name)}'`) {
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
      if ((opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(id)}`) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: newName, description: description })
      ) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        teamId: teamId,
        name: name,
        newName: newName,
        description: description
      }
    });
  });

  it('correctly patches channel updates by teamName and id', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(teamName)}'`) {
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
      if ((opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(id)}`) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: newName, description: description })
      ) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        debug: false,
        teamName: teamName,
        id: id,
        newName: newName,
        description: description
      }
    });
  });

  it('fails when team name does not exist', async () => {
    const errorMessage = 'The specified team does not exist in the Microsoft Teams';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent(teamName)}'`) {
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
        confirm: true
      }
    }), new CommandError(errorMessage));
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