import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../cli/cli.js';
import request from '../request.js';
import { sinonUtil } from './sinonUtil.js';
import { customAppScope } from './customAppScope.js';
import { formatting } from './formatting.js';
import { settingsNames } from '../settingsNames.js';

describe('utils/customAppScope', () => {
  const displayName = 'Users from Marketing Department';
  const invalidDisplayName = 'User from Marketing Deprtment';
  const customAppScopeResponse = {
    "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
    "type": "RecipientScope",
    "displayName": "Users from Marketing Department",
    "customAttributes": {
      "Exclusive": false,
      "RecipientFilter": "Department -eq 'Marketing'"
    }
  };
  const customAppScopeLimitedResponse = {
    "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
    "displayName": "Users from Marketing Department"
  };
  const secondCustomAppScopeResponse = {
    "id": "8044456e-ef79-4a2b-97d4-d7605acbde76",
    "type": "RecipientScope",
    "displayName": "Managers",
    "customAttributes": {
      "Exclusive": false,
      "RecipientFilter": "Title -like '*Manager*'"
    }
  };

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get single custom application scope by name using getCustomAppScopeByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            customAppScopeResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await customAppScope.getCustomAppScopeByDisplayName(displayName);
    assert.deepStrictEqual(actual, {
      "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
      "type": "RecipientScope",
      "displayName": "Users from Marketing Department",
      "customAttributes": {
        "Exclusive": false,
        "RecipientFilter": "Department -eq 'Marketing'"
      }
    });
  });

  it('correctly get single custom application scope by name using getCustomAppScopeByDisplayName with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'&$select=id,displayName`) {
        return {
          value: [
            customAppScopeLimitedResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    const actual = await customAppScope.getCustomAppScopeByDisplayName(displayName, 'id,displayName');
    assert.deepStrictEqual(actual, {
      "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
      "displayName": "Users from Marketing Department"
    });
  });

  it('handles selecting single custom application scope when multiple custom application scopes with the specified name found using getCustomAppScopeByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            customAppScopeResponse,
            secondCustomAppScopeResponse
          ]
        };
      }

      throw 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves(customAppScopeResponse);

    const actual = await customAppScope.getCustomAppScopeByDisplayName(displayName);
    assert.deepStrictEqual(actual, {
      "id": "0d2ba203-f407-4de0-9f6a-7e411484e4da",
      "type": "RecipientScope",
      "displayName": "Users from Marketing Department",
      "customAttributes": {
        "Exclusive": false,
        "RecipientFilter": "Department -eq 'Marketing'"
      }
    });
  });

  it('throws error message when no custom application scope was found using getCustomAppScopeByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(invalidDisplayName)}'`) {
        return { value: [] };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(customAppScope.getCustomAppScopeByDisplayName(invalidDisplayName)), Error(`The specified custom application scope '${invalidDisplayName}' does not exist.`);
  });

  it('throws error message when multiple custom application scopes were found using getCustomAppScopeByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/roleManagement/exchange/customAppScopes?$filter=displayName eq '${formatting.encodeQueryParameter(displayName)}'`) {
        return {
          value: [
            customAppScopeResponse,
            secondCustomAppScopeResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(customAppScope.getCustomAppScopeByDisplayName(displayName),
      Error(`Multiple custom application scopes with name '${displayName}' found. Found: ${customAppScopeResponse.id}, ${secondCustomAppScopeResponse.id}.`));
  });
});