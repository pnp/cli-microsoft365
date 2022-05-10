import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import { planner } from './planner';
import { sinonUtil } from "./sinonUtil";

const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
const validPlanName = 'Plan name';
const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';

const singlePlanResponse = {
  id: validPlanId,
  title: validPlanName,
  owner: validOwnerGroupId
};

describe('utils/planner', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  it('correctly get a single plan by id.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        return singlePlanResponse;
      }

      return 'Invalid Request';
    });

    const actual = await planner.getPlanById(validPlanId);
    assert.strictEqual(actual, singlePlanResponse);
  });

  it('display error message when plan is not found.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${validPlanId}`) {
        throw Error('Plan not found.');
      }

      return 'Invalid Request';
    });

    try {
      await planner.getPlanById(validPlanId);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Planner plan with id ${validPlanId} was not found.`));
    }
  });
});