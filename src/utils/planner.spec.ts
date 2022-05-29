import * as assert from 'assert';
import * as sinon from 'sinon';
import request from "../request";
import { PlannerPlan } from '@microsoft/microsoft-graph-types';
import { planner } from './planner';
import { sinonUtil } from "./sinonUtil";

const validPlanId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
const validPlanTitle = 'Plan title';
const validOwnerGroupId = '00000000-0000-0000-0000-000000000000';

const singlePlanResponse: PlannerPlan = {
  id: validPlanId,
  title: validPlanTitle,
  owner: validOwnerGroupId
};

const multiplePlanResponse = {
  value: [
    singlePlanResponse
  ] as PlannerPlan[]
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
        return multiplePlanResponse;
      }

      return 'Invalid Request';
    });

    const actual = await planner.getPlansByGroupId(validOwnerGroupId);
    assert.strictEqual(actual, multiplePlanResponse.value);
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
      assert.deepStrictEqual(ex, Error(`Planner plan with id '${validPlanId}' was not found.`));
    }
  });

  it('correctly get plan by title.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return multiplePlanResponse;
      }

      return 'Invalid Request';
    });

    const actual = await planner.getPlanByTitle(validPlanTitle, validOwnerGroupId);
    assert.strictEqual(actual, singlePlanResponse);
  });

  it('fails to get plan when plan doesn not exist', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        const response = {...multiplePlanResponse};
        response.value[0].title = "Wrong title";
        return response;
      }

      return 'Invalid Request';
    });

    try {
      await planner.getPlanByTitle(validPlanTitle, validOwnerGroupId);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`The specified plan '${validPlanTitle}' does not exist.`));
    }
  });

  it('fails to get plan when multiple plans have the same title', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${validOwnerGroupId}/planner/plans`) {
        return {
          value: [
            { title: validPlanTitle, id: validPlanId },
            { title: validPlanTitle, id: validPlanId }
          ]
        };
      }

      return 'Invalid Request';
    });

    try {
      await planner.getPlanByTitle(validPlanTitle, validOwnerGroupId);
      assert.fail('No error message thrown.');
    }
    catch (ex) {
      assert.deepStrictEqual(ex, Error(`Multiple plans with title '${validPlanTitle}' found: ${[validPlanId, validPlanId]}.`));
    }
  });
});