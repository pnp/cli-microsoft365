import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import { aadGroup } from './aadGroup';
import { formatting } from './formatting';
import { sinonUtil } from "./sinonUtil";

const validGroupName = 'Group name';
const validGroupId = '00000000-0000-0000-0000-000000000000';

const singleGroupResponse = {
  id: validGroupId,
  displayName: validGroupName
};

describe('utils/aadGroup', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get
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

    try {
      await aadGroup.getGroupByDisplayName(validGroupName);
      assert.fail('Error expected, but was not thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified group '${validGroupName}' does not exist.`));
    }
  });

  it('throws error message when multiple groups were found using getGroupByDisplayName', async () => {
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

    try {
      await aadGroup.getGroupByDisplayName(validGroupName);
      assert.fail('Error expected, but was not thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple groups with name '${validGroupName}' found: ${[validGroupId, validGroupId]}.`));
    }
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'?$select=id`) {
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
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'?$select=id`) {
        return { value: [] };
      }

      return 'Invalid Request';
    });

    try {
      await aadGroup.getGroupIdByDisplayName(validGroupName);
      assert.fail('Error expected, but was not thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified group '${validGroupName}' does not exist.`));
    }
  });

  it('throws error message when multiple groups were found using getGroupIdByDisplayName', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter(validGroupName)}'?$select=id`) {
        return {
          value: [
            { id: validGroupId },
            { id: validGroupId }
          ]
        };
      }

      return 'Invalid Request';
    });

    try {
      await aadGroup.getGroupIdByDisplayName(validGroupName);
      assert.fail('Error expected, but was not thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple groups with name '${validGroupName}' found: ${[validGroupId, validGroupId]}.`));
    }
  });
}); 