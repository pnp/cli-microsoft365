import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import { aadGroup } from './aadGroup';
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

  it('display error message when group is not found.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validGroupId}`) {
        throw Error('Group not found.');
      }

      return 'Invalid Request';
    });

    try {
      await aadGroup.getGroupById(validGroupId);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Group with ID ${validGroupId} was not found.`));
    }
  });
}); 