import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import { aadGroup } from './aadGroup.js';
import { formatting } from './formatting.js';
import { sinonUtil } from "./sinonUtil.js";
import { Logger } from '../cli/Logger.js';
import { Cli } from '../cli/Cli.js';
import { settingsNames } from '../settingsNames.js';

const validGroupName = 'Group name';
const validGroupId = '00000000-0000-0000-0000-000000000000';

const singleGroupResponse = {
  id: validGroupId,
  displayName: validGroupName
};

describe('utils/aadGroup', () => {
  let cli: Cli;
  let logger: Logger;
  let log: string[];

  before(() => {
    cli = Cli.getInstance();
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  it('correctly get a single group by id.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return singleGroupResponse;
      }

      return 'Invalid Request';
    });

    const actual = await aadGroup.getGroupById(validGroupId);
    assert.strictEqual(actual, singleGroupResponse);
  });

  it('throws error message when no group was found using getGroupByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(aadGroup.getGroupByDisplayName(validGroupName), Error(`The specified group '${validGroupName}' does not exist.`));
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

    await assert.rejects(aadGroup.getGroupByDisplayName(validGroupName), Error("Multiple groups with name 'Group name' found. Found: 00000000-0000-0000-0000-000000000000."));
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

    const actual = await aadGroup.getGroupByDisplayName(validGroupName);
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

    const actual = await aadGroup.getGroupIdByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, singleGroupResponse.id);
  });

  it('throws error message when no group was found using getGroupIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'&$select=id`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    await assert.rejects(aadGroup.getGroupIdByDisplayName(validGroupName), Error(`The specified group '${validGroupName}' does not exist.`));
  });

  it('throws error message when multiple groups were found using getGroupIdByDisplayName', async () => {
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

    await assert.rejects(aadGroup.getGroupIdByDisplayName(validGroupName), Error(`Multiple groups with name '${validGroupName}' found: ${[validGroupId, validGroupId]}.`));
  });

  it('updates a group to public successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await aadGroup.setGroup(validGroupId, true, logger, true);
    assert(patchStub.called);
  });

  it('updates a group to private successfully', async () => {
    const patchStub = sinon.stub(request, 'patch').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        return;
      }

      return 'Invalid Request';
    });

    await aadGroup.setGroup(validGroupId, false, logger, true);
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

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves({ id: validGroupId, displayName: validGroupName });

    const actual = await aadGroup.getGroupByDisplayName(validGroupName);
    assert.deepStrictEqual(actual, singleGroupResponse);
  });
}); 