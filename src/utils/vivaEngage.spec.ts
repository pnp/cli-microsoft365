import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { vivaEngage } from './vivaEngage.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/vivaEngage', () => {
  const displayName = 'All Company';
  const invalidDisplayName = 'All Compayn';
  const entraGroupId = '0bed8b86-5026-4a93-ac7d-56750cc099f1';
  const communityId = 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9';
  const communityResponse = {
    "id": "eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9",
    "description": "This is the default group for everyone in the network",
    "displayName": "All Company",
    "privacy": "Public",
    "groupId": "0bed8b86-5026-4a93-ac7d-56750cc099f1"
  };
  const anotherCommunityResponse = {
    "id": "eyJfdHlwZ0NzY5SIwiIiSJ9IwO6IaWQiOIMTM1ODikdyb3Vw",
    "description": "Test only",
    "displayName": "All Company",
    "privacy": "Private",
    "groupId": "0bed8b86-5026-4a93-ac7d-56750cc099f1"
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single community id by name using getCommunityIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            communityResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await vivaEngage.getCommunityIdByDisplayName(displayName);
    assert.deepStrictEqual(actual, 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9');
  });

  it('handles selecting single community when multiple communities with the specified name found using getCommunityIdByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            communityResponse,
            anotherCommunityResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(communityResponse);

    const actual = await vivaEngage.getCommunityIdByDisplayName(displayName);
    assert.deepStrictEqual(actual, 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9');
  });

  it('throws error message when no community was found using getCommunityIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(vivaEngage.getCommunityIdByDisplayName(invalidDisplayName)), Error(`The specified Viva Engage community '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple communities were found using getCommunityIdByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            communityResponse,
            anotherCommunityResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(vivaEngage.getCommunityIdByDisplayName(displayName),
      Error(`Multiple Viva Engage communities with name '${displayName}' found. Found: ${communityResponse.id}, ${anotherCommunityResponse.id}.`));
  });

  it('correctly get single community id by group id using getCommunityIdByEntraGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/employeeExperience/communities?$select=id,groupId') {
        return {
          value: [
            communityResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await vivaEngage.getCommunityIdByEntraGroupId(entraGroupId);
    assert.deepStrictEqual(actual, 'eyJfdHlwZSI6Ikdyb3VwIiwiaWQiOiI0NzY5MTM1ODIwOSJ9');
  });

  it('throws error message when no community was found using getCommunityIdByEntraGroupId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/employeeExperience/communities?$select=id,groupId') {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(vivaEngage.getCommunityIdByEntraGroupId(entraGroupId)), Error(`The Microsoft Entra group with id '${entraGroupId}' is not associated with any Viva Engage community.`);
  });

  it('correctly gets Entra group ID by community ID using getEntraGroupIdByCommunityId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}?$select=groupId`) {
        return communityResponse;
      }

      throw 'Invalid Request';
    });

    const actual = await vivaEngage.getEntraGroupIdByCommunityId(communityId);
    assert.deepStrictEqual(actual, '0bed8b86-5026-4a93-ac7d-56750cc099f1');
  });

  it('throws error message when no Entra group ID was found using getEntraGroupIdByCommunityId', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}?$select=groupId`) {
        return null;
      }

      throw 'Invalid Request';
    });

    await assert.rejects(vivaEngage.getEntraGroupIdByCommunityId(communityId)), Error(`The specified Viva Engage community with ID '${communityId}' does not exist.`);
  });

  it('correctly gets Entra group ID by community display name using getEntraGroupIdByCommunityDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities/${communityId}?$select=groupId`) {
        return communityResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/employeeExperience/communities?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            communityResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await vivaEngage.getEntraGroupIdByCommunityDisplayName(displayName);
    assert.deepStrictEqual(actual, entraGroupId);
  });
});