import assert = require('assert');
import * as sinon from 'sinon';
import request from "../request";
import { planner } from './planner';
import { sinonUtil } from "./sinonUtil";

const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
const validPlanName = 'Plan name';
const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';

const singlePlanResponse = {
  "value": [
    {
      "id": validPlanId,
      "title": validPlanName,
      "owner": validOwnerGroupId
    }
  ]
};

describe('utils/planner', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  it('correctly get all plans related to a specific group.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return singlePlanResponse;
      }

      return 'Invalid Request';
    });

    const actual = await planner.getPlansByGroupId(validOwnerGroupId);
    assert.strictEqual(actual, singlePlanResponse.value);
  });
});