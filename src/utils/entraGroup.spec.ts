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
      request.post,
      cli.getSettingWithDefaultValue,
      cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get a single group by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return singleGroupResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupById(validGroupId);
    assert.strictEqual(actual, singleGroupResponse);
  });

  it('correctly get a single group by id with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}?$select=id,displayName`) {
        return singleGroupResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraGroup.getGroupById(validGroupId, 'id,displayName');
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

  it('updates a group to private successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await entraGroup.setGroup(validGroupId, true, 'display name', 'description', logger, true);
    assert(patchStub.called);
  });

  it('updates a group to public successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await entraGroup.setGroup(validGroupId, false, 'display name', 'description', logger, true);
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

  it('handles selecting single result when multiple groups with the specified name found and cli is set to prompt with specified properties', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id,displayName`) {
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

    const actual = await entraGroup.getGroupByDisplayName(validGroupName, 'id,displayName');
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

  it('correctly gets group ids by display names', async () => {
    const groupNames = ['group1', 'group2', 'group3', 'group4', 'group5', 'group6', 'group7', 'group8', 'group9', 'group10', 'group11', 'group12', 'group13', 'group14', 'group15', 'group16', 'group17', 'group18', 'group19', 'group20', 'group21', 'group22', 'group23', 'group24', 'group25'];
    const groupIds = ['5acd04e0-1234-abcd-0a9c-000000000001', '5acd04e0-1234-abcd-0a9c-000000000002', '5acd04e0-1234-abcd-0a9c-000000000003', '5acd04e0-1234-abcd-0a9c-000000000004', '5acd04e0-1234-abcd-0a9c-000000000005', '5acd04e0-1234-abcd-0a9c-000000000006', '5acd04e0-1234-abcd-0a9c-000000000007', '5acd04e0-1234-abcd-0a9c-000000000008', '5acd04e0-1234-abcd-0a9c-000000000009', '5acd04e0-1234-abcd-0a9c-000000000010', '5acd04e0-1234-abcd-0a9c-000000000011', '5acd04e0-1234-abcd-0a9c-000000000012', '5acd04e0-1234-abcd-0a9c-000000000013', '5acd04e0-1234-abcd-0a9c-000000000014', '5acd04e0-1234-abcd-0a9c-000000000015', '5acd04e0-1234-abcd-0a9c-000000000016', '5acd04e0-1234-abcd-0a9c-000000000017', '5acd04e0-1234-abcd-0a9c-000000000018', '5acd04e0-1234-abcd-0a9c-000000000019', '5acd04e0-1234-abcd-0a9c-000000000020', '5acd04e0-1234-abcd-0a9c-000000000021', '5acd04e0-1234-abcd-0a9c-000000000022', '5acd04e0-1234-abcd-0a9c-000000000023', '5acd04e0-1234-abcd-0a9c-000000000024', '5acd04e0-1234-abcd-0a9c-000000000025'];

    let batch = -1;
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: groupIds.slice(++batch * 20, batch * 20 + 20).map(groupId => ({
            status: 200,
            body: {
              value: [{
                id: groupId
              }]
            }
          }))
        };
      }

      throw 'Invalid request';
    });

    const actual = await entraGroup.getGroupIdsByDisplayNames(groupNames);
    assert.deepStrictEqual(postStub.firstCall.args[0].data.requests, groupNames.slice(0, 20).map((name, i) => ({ id: i + 1, method: 'GET', url: `/groups?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'&$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, groupNames.slice(20, 40).map((name, i) => ({ id: i + 1, method: 'GET', url: `/groups?$filter=displayName eq '${formatting.encodeQueryParameter(name)}'&$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(actual, groupIds);
  });

  it('correctly throws error when no group was found with a specific display name', async () => {
    const groupNames = ['group1', 'group2', 'group3', 'group4', 'group5', 'group6', 'group7', 'group8', 'group9', 'group10', 'group11', 'group12', 'group13', 'group14', 'group15', 'group16', 'group17', 'group18', 'group19', 'group20', 'group21', 'group22', 'group23', 'group24', 'group25'];
    const groupIds = ['5acd04e0-1234-abcd-0a9c-000000000001', '5acd04e0-1234-abcd-0a9c-000000000002', '5acd04e0-1234-abcd-0a9c-000000000003', '5acd04e0-1234-abcd-0a9c-000000000004', '5acd04e0-1234-abcd-0a9c-000000000005', '5acd04e0-1234-abcd-0a9c-000000000006', '5acd04e0-1234-abcd-0a9c-000000000007', '5acd04e0-1234-abcd-0a9c-000000000008', '5acd04e0-1234-abcd-0a9c-000000000009', '5acd04e0-1234-abcd-0a9c-000000000010', '5acd04e0-1234-abcd-0a9c-000000000011', '5acd04e0-1234-abcd-0a9c-000000000012', '5acd04e0-1234-abcd-0a9c-000000000013', '5acd04e0-1234-abcd-0a9c-000000000014', '5acd04e0-1234-abcd-0a9c-000000000015', '5acd04e0-1234-abcd-0a9c-000000000016', '5acd04e0-1234-abcd-0a9c-000000000017', '5acd04e0-1234-abcd-0a9c-000000000018', '5acd04e0-1234-abcd-0a9c-000000000019', '5acd04e0-1234-abcd-0a9c-000000000020', '5acd04e0-1234-abcd-0a9c-000000000021', '5acd04e0-1234-abcd-0a9c-000000000022', '5acd04e0-1234-abcd-0a9c-000000000023', '5acd04e0-1234-abcd-0a9c-000000000024', '5acd04e0-1234-abcd-0a9c-000000000025'];

    let counter = 0;
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: groupIds.slice(counter, counter + 20).map(groupId => {
            if (counter++ < groupNames.length - 1) {
              return {
                status: 200,
                body: {
                  value: [{
                    id: groupId
                  }]
                }
              };
            }
            else {
              return {
                id: counter % 20,
                status: 200,
                body: {
                  value: []
                }
              };
            }
          })
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(entraGroup.getGroupIdsByDisplayNames(groupNames), Error(`The specified group with name '${groupNames[groupNames.length - 1]}' does not exist.`));
  });

  it('correctly throws error when multiple groups were found with a specific display name', async () => {
    const groupNames = ['group1', 'group2', 'group3', 'group4', 'group5', 'group6', 'group7', 'group8', 'group9', 'group10', 'group11', 'group12', 'group13', 'group14', 'group15', 'group16', 'group17', 'group18', 'group19', 'group20', 'group21', 'group22', 'group23', 'group24', 'group25'];
    const groupIds = ['5acd04e0-1234-abcd-0a9c-000000000001', '5acd04e0-1234-abcd-0a9c-000000000002', '5acd04e0-1234-abcd-0a9c-000000000003', '5acd04e0-1234-abcd-0a9c-000000000004', '5acd04e0-1234-abcd-0a9c-000000000005', '5acd04e0-1234-abcd-0a9c-000000000006', '5acd04e0-1234-abcd-0a9c-000000000007', '5acd04e0-1234-abcd-0a9c-000000000008', '5acd04e0-1234-abcd-0a9c-000000000009', '5acd04e0-1234-abcd-0a9c-000000000010', '5acd04e0-1234-abcd-0a9c-000000000011', '5acd04e0-1234-abcd-0a9c-000000000012', '5acd04e0-1234-abcd-0a9c-000000000013', '5acd04e0-1234-abcd-0a9c-000000000014', '5acd04e0-1234-abcd-0a9c-000000000015', '5acd04e0-1234-abcd-0a9c-000000000016', '5acd04e0-1234-abcd-0a9c-000000000017', '5acd04e0-1234-abcd-0a9c-000000000018', '5acd04e0-1234-abcd-0a9c-000000000019', '5acd04e0-1234-abcd-0a9c-000000000020', '5acd04e0-1234-abcd-0a9c-000000000021', '5acd04e0-1234-abcd-0a9c-000000000022', '5acd04e0-1234-abcd-0a9c-000000000023', '5acd04e0-1234-abcd-0a9c-000000000024', '5acd04e0-1234-abcd-0a9c-000000000025'];

    let counter = 0;
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: groupIds.slice(counter, counter + 20).map(groupId => {
            if (counter++ < groupNames.length - 1) {
              return {
                status: 200,
                body: {
                  value: [{
                    id: groupId
                  }]
                }
              };
            }
            else {
              return {
                id: counter % 20,
                status: 200,
                body: {
                  value: [{
                    id: groupId
                  },
                  {
                    id: '5acd04e0-1234-abcd-0a9c-000000000026'
                  }]
                }
              };
            }
          })
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(entraGroup.getGroupIdsByDisplayNames(groupNames), Error(`Multiple groups with the name '${groupNames[groupNames.length - 1]}' found.`));
  });
}); 