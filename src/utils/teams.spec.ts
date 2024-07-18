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
  before(() => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => settingName === settingsNames.prompt ? false : defaultValue);
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      //cli.getSettingWithDefaultValue,
      teams.getTeamIdByDisplayName,
      teams.getChannelIdByDisplayName
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('correctly gets team id by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse] };
      }

      throw 'Invalid Request';
    });

    const actual = await teams.getTeamIdByDisplayName(teamName);
    assert.strictEqual(actual, teamId);
  });

  it('throws error if no team is found when retrieving teamId by displayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.getTeamIdByDisplayName(teamName), Error(`The specified team '${teamName}' does not exist.`));
  });

  it('throws error message when multiple teams were found using getTeamIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse, { id: 'df20c966-aa55-4810-a086-7e20001e0788' }] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(teams.getTeamIdByDisplayName(teamName), Error(`Multiple teams with name '${teamName}' found. Found: ${teamId}, df20c966-aa55-4810-a086-7e20001e0788.`));
  });

  it('handles selecting single result when multiple teams with the specified name found using getTeamIdByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams?$filter=displayName eq '${formatting.encodeQueryParameter(teamName)}'&$select=id`) {
        return { value: [teamResponse, { id: 'df20c966-aa55-4810-a086-7e20001e0788' }] };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: teamId });

    const actual = await teams.getTeamIdByDisplayName(teamName);
    assert.deepStrictEqual(actual, teamId);
  });

  it('correctly retrieves channel id by displayName', async () => {
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

    await assert.rejects(teams.getChannelIdByDisplayName(teamId, channelName), Error(`The specified channel ${channelName} does not exist in the Microsoft Teams team ${teamId}.`));
  });

}); 
