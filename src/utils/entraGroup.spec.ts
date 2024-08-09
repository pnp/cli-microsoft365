import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import { entraGroup } from './entraGroup.js';
import { formatting } from './formatting.js';
import { sinonUtil } from "./sinonUtil.js";
import { Logger } from '../cli/Logger.js';
import { cli } from '../cli/cli.js';
import { settingsNames } from '../settingsNames.js';

const validGroupName = 'Group name';
const validGroupId = '00000000-0000-0000-0000-000000000000';
const validGroupMailNickname = 'GroupName';

const singleGroupResponse = {
  id: validGroupId,
  displayName: validGroupName
};

describe('utils/entraGroup', () => {
  let logger: Logger;
  let log: string[];

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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get a single group by id.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return singleGroupResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupById(validGroupId);
    assert.strictEqual(actual, singleGroupResponse);
  });

  it('throws error message when no group was found using getGroupByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupByDisplayName(validGroupName), Error(`The specified group '${validGroupName}' does not exist.`));
  });

  it('throws error message when multiple groups were found using getGroupByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'`) {
        return {
          value: [
            { id: validGroupId, displayName: validGroupName },
            { id: validGroupId, displayName: validGroupName }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupByDisplayName(validGroupName), Error("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
  });

  it('correctly get single group by name using getGroupByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'`) {
        return {
          value: [
            singleGroupResponse
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, singleGroupResponse);
  });

  it('correctly get single group id by name using getGroupIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return {
          value: [
            { id: singleGroupResponse.id }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupIdByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, singleGroupResponse.id);
  });

  it('throws error message when no group was found using getGroupIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupIdByDisplayName(validGroupName), Error(`The specified group '${validGroupName}' does not exist.`));
  });

  it('throws error message when multiple groups were found using getGroupIdByDisplayName', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: validGroupId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupIdByDisplayName(validGroupName), Error(`Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000.`));
  });

  it('handles selecting single result when multiple groups with the specified name found using getGroupIdByDisplayName and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: validGroupId }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: validGroupId });

    const actual = await entraGroup.getGroupIdByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, validGroupId);
  });

  it('correctly get single group id by name using getGroupIdByMailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'&$select=id`) {
        return {
          value: [
            { id: singleGroupResponse.id }
          ]
        };
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupIdByMailNickname(validGroupMailNickname);
    assert.deepStrictEqual(actual, singleGroupResponse.id);
  });

  it('throws error message when no group was found using getGroupIdByMailNickname', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'&$select=id`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupIdByMailNickname(validGroupMailNickname), Error(`The specified group '${validGroupMailNickname}' does not exist.`));
  });

  it('throws error message when multiple groups were found using getGroupIdByMailNickname', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'&$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: validGroupId }
          ]
        };
      }

      return 'Invalid Request';
    });

    await assert.rejects(entraGroup.getGroupIdByMailNickname(validGroupMailNickname), Error(`Multiple groups with mail nickname 'GroupName' found. Found: 00000000-0000-0000-0000-000000000000.`));
  });

  it('handles selecting single result when multiple groups with the specified name found using getGroupIdByMailNickname and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=mailNickname eq '${formatting.encodeQueryParameter(validGroupMailNickname)}'&$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: validGroupId }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: validGroupId });

    const actual = await entraGroup.getGroupIdByMailNickname(validGroupMailNickname);
    assert.deepStrictEqual(actual, validGroupId);
  });

  it('updates a group to public successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await entraGroup.setGroup(validGroupId, true, logger, true);
    assert(patchStub.called);
  });

  it('updates a group to private successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await entraGroup.setGroup(validGroupId, false, logger, true);
    assert(patchStub.called);
  });

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'`) {
        return {
          value: [
            { id: validGroupId, displayName: validGroupName },
            { id: validGroupId, displayName: validGroupName }
          ]
        };
      }

      return 'Invalid Request';
    });

    sinon.stub(cli, 'handleMultipleResultsFound').resolves({ id: validGroupId, displayName: validGroupName });

    const actual = await entraGroup.getGroupByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, singleGroupResponse);
  });

  it('returns true if group is a valid m365group', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}?$select=groupTypes`) {
        return {
          groupTypes: [
            'Unified'
          ]
        };
      }

      return 'Invalid Request';
    });
    const actual = await entraGroup.isUnifiedGroup(validGroupId);
    assert.deepStrictEqual(actual, true);
  });

  it('returns false if group is not a m365group', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}?$select=groupTypes`) {
        return {
          groupTypes: []
        };
      }

      return 'Invalid Request';
    });
    const actual = await entraGroup.isUnifiedGroup(validGroupId);
    assert.deepStrictEqual(actual, false);
  });
}); 