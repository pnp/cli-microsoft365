import sinon from 'sinon';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { formatting } from './formatting.js';
import { teams } from './teams.js';
import assert from 'assert';
import { cli } from '../cli/cli.js';
import { settingsNames } from '../settingsNames.js';

const teamName = 'HR Team';
const teamId = '0b0b204f-7ca0-4c7f-baf2-53caa381828b';
const teamResponse = {
  id: teamId
};
const channelName = 'General';
const channelId = '19:7a3a82caa8f8436889fbb017acdb11b6@thread.tacv2';
const channelResponse = { id: channelId };

describe('utils/teams', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      teams.getTeamIdByDisplayName,
      teams.verifyChannelExistsById,
      teams.getChannelIdByDisplayName,
      teams.getTeamId,
      teams.verifyTeamExistsById
    ]);
  });

  it('gets team id by displayName', async () => {
    sinon.stub(teams, 'getTeamIdByDisplayName').resolves(teamId);

    const actual = await teams.getTeamId(undefined, teamName);
    assert.strictEqual(actual, teamId);
  });

  it('returns team id and verifies that team exists', async () => {
    sinon.stub(teams, 'verifyTeamExistsById').resolves(teamId);

    const actual = await teams.getTeamId(teamId, undefined);
    assert.strictEqual(actual, teamId);
  });

  it('verifies that team exists by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}?$select=id`) {
        return { id: teamId };
      }

      throw 'Invalid Request';
    });

    const actual = await teams.verifyTeamExistsById(teamId);
    assert.strictEqual(actual, teamId);
  });

  it('correctly get team id by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse] };
      }

      throw 'Invalid Request';
    });

    const actual = await teams.getTeamIdByDisplayName(teamName);
    assert.strictEqual(actual, teamId);
  });

  it('throws error if no teams are found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.getTeamIdByDisplayName(teamName), Error(`The specified team '${teamName}' does not exist.`));
  });

  it('throws error message when multiple teams were found using getTeamIdByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse, teamResponse] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.getTeamIdByDisplayName(teamName), Error(`Multiple teams with name '${teamName}' found. Found: ${teamId}.`));
  });

  it('handles selecting single result when multiple teams with the specified name found using getTeamIdByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse, teamResponse] };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: teamId });

    const actual = await teams.getTeamIdByDisplayName(teamName);
    assert.deepStrictEqual(actual, teamId);
  });

  it('gets channel id by name', async () => {
    sinon.stub(teams, 'getChannelIdByDisplayName').resolves(channelId);

    const actual = await teams.getChannelId(teamId, undefined, channelName);
    assert.strictEqual(actual, channelId);
  });

  it('returns channel id and verifies that channel exists', async () => {
    sinon.stub(teams, 'verifyChannelExistsById').resolves(channelId);

    const actual = await teams.getChannelId(teamId, channelId, undefined);
    assert.strictEqual(actual, channelId);
  });


  it('verifies that channel exists by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}?$select=id`) {
        return { id: channelId };
      }

      throw 'Invalid Request';
    });

    const actual = await teams.verifyChannelExistsById(teamId, channelId);
    assert.strictEqual(actual, channelId);
  });

  it('throws error when channel does not exist by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}?$select=id`) {
        throw {
          error: {
            code: 'NotFound',
            message: 'NotFound',
            innerError: {
              code: '1',
              message: `LocationLookupFailed-Location lookup failed for thread ${channelId}`,
              date: '2024-02-21T20:08:18',
              'request-id': 'f32dfea8-1a1b-4c4b-8610-ada8b9c10a84',
              'client-request-id': 'f32dfea8-1a1b-4c4b-8610-ada8b9c10a84'
            }
          }
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.verifyChannelExistsById(teamId, channelId), Error('The specified channel does not exist in the Microsoft Teams team.'));
  });

  it('correctly get channel id by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(channelName)}'&$select=id`) {
        return { value: [channelResponse] };
      }

      throw 'Invalid Request';
    });

    const actual = await teams.getChannelIdByDisplayName(teamId, channelName);
    assert.strictEqual(actual, channelId);
  });

  it('throws error if no channel with the specified name is found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${teamId}/channels?$filter=displayName eq '${formatting.encodeQueryParameter(channelName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.getChannelIdByDisplayName(teamId, channelName), Error('The specified channel does not exist in the Microsoft Teams team'));
  });

}); 
