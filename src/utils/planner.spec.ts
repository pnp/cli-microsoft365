import assert from 'assert';
import sinon from 'sinon';
import request from "../request.js";
import { PlannerBucket, PlannerPlan } from '@microsoft/microsoft-graph-types';
import { planner } from './planner.js';
import { sinonUtil } from "./sinonUtil.js";
import { cli } from '../cli/cli.js';
import { formatting } from './formatting.js';

const planId = 'oUHpnKBFekqfGE_PS6GGUZcAFY7b';
const planTitle = 'Plan title';
const ownerGroupId = '00000000-0000-0000-0000-000000000000';
const rosterId = 'tYqYlNd6eECmsNhN_fcq85cAGAnd';
const bucketName = 'To Do';
const bucketId = 'iTpUXFFYlkyyUsoWmILQbJgAMCbR';

const singlePlanResponse: PlannerPlan = {
  id: planId,
  title: planTitle,
  container: {
    containerId: ownerGroupId
  }
};

const singlePlanResponseTrimmed: PlannerPlan = {
  id: planId,
  title: planTitle
};

const multiplePlanResponse = {
  value: [
    {
      id: 'RnZkyGveikyfG2u817hodJgADNwq',
      title: 'Another plan',
      container: {
        containerId: '00000000-0000-0000-0000-000000000000'
      }
    },
    singlePlanResponse
  ] as PlannerPlan[]
};

const multiplePlanResponseTrimmed = {
  value: multiplePlanResponse.value.map(p => ({ id: p.id, title: p.title }))
};

const singleBucketResponse: PlannerBucket = {
  name: bucketName,
  planId: planId,
  orderHint: '8584914112[Y',
  id: bucketId
};

const multipleBucketResponse = {
  value: [
    {
      name: 'In Progress',
      planId: planId,
      orderHint: '8584914113862517879P=',
      id: 'X5vwdkKh70-5hsFT1MJ8-pgAL4Si'
    },
    singleBucketResponse
  ]
};

const multipleBucketResponseTrimmed = {
  value: multipleBucketResponse.value.map(b => ({ name: b.name, id: b.id }))
};

describe('utils/planner', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get,
      cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('correctly get all plans related to a specific group', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return multiplePlanResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlansByGroupId(ownerGroupId);
    assert.deepStrictEqual(actual, multiplePlanResponse.value);
  });

  it('correctly get the plan from a roster', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${rosterId}/plans`) {
        return { value: [singlePlanResponse] };
      }

      return 'Invalid Request: ' + opts.url;
    });

    const actual = await planner.getPlanByRosterId(rosterId);
    assert.deepStrictEqual(actual, singlePlanResponse);
  });

  it('correctly throws an error when a roster does not have a plan', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${rosterId}/plans`) {
        return { value: [] };
      }

      return 'Invalid Request: ' + opts.url;
    });

    await assert.rejects(planner.getPlanByRosterId(rosterId), Error(`The specified roster '${rosterId}' does not have a plan.`));
  });

  it('correctly get the plan ID from a roster', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${rosterId}/plans?$select=id`) {
        return { value: [{ id: planId }] };
      }

      return 'Invalid Request: ' + opts.url;
    });

    const actual = await planner.getPlanIdByRosterId(rosterId);
    assert.deepStrictEqual(actual, planId);
  });

  it('correctly throws an error when a roster does not have a plan when getting the plan ID', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/beta/planner/rosters/${rosterId}/plans?$select=id`) {
        return { value: [] };
      }

      return 'Invalid Request: ' + opts.url;
    });

    await assert.rejects(planner.getPlanIdByRosterId(rosterId), Error(`The specified roster '${rosterId}' does not have a plan.`));
  });

  it('correctly get a single plan by id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}`) {
        return singlePlanResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlanById(planId);
    assert.deepStrictEqual(actual, singlePlanResponse);
  });

  it('display error message when plan is not found', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}`) {
        throw Error('Plan not found.');
      }

      return 'Invalid Request ' + opts.url;
    });

    await assert.rejects(planner.getPlanById(planId), Error(`Planner plan with id '${planId}' was not found.`));
  });

  it('correctly get plan by title', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return multiplePlanResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlanByTitle(planTitle, ownerGroupId);
    assert.deepStrictEqual(actual, singlePlanResponse);
  });

  it('fails to get plan when plan does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return multiplePlanResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    await assert.rejects(planner.getPlanByTitle('Wrong title', ownerGroupId),
      Error(`The specified plan 'Wrong title' does not exist.`));
  });

  it('correctly returns a plan when multiple plans have the same title with prompt', async () => {
    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(singlePlanResponse);

    const requestResult = [
      singlePlanResponse,
      { ...singlePlanResponse, id: 'RnZkyGveikyfG2u817hodJgADNwq' }
    ];

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans`) {
        return {
          value: requestResult
        };
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlanByTitle(planTitle, ownerGroupId);

    const plansKeyValuePair = formatting.convertArrayToHashTable('id', requestResult);
    assert(stubMultiResults.calledOnceWithExactly(`Multiple plans with title '${planTitle}' found.`, plansKeyValuePair));
    assert.deepStrictEqual(actual, singlePlanResponse);
  });

  it('correctly get plan ID by title.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans?$select=id,title`) {
        return multiplePlanResponseTrimmed;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlanIdByTitle(planTitle, ownerGroupId);
    assert.deepStrictEqual(actual, planId);
  });

  it('fails to get plan ID when plan does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans?$select=id,title`) {
        return multiplePlanResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    await assert.rejects(planner.getPlanIdByTitle('Wrong title', ownerGroupId),
      Error(`The specified plan 'Wrong title' does not exist.`));
  });

  it('correctly returns a plan ID when multiple plans have the same title with prompt', async () => {
    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(singlePlanResponse);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/${ownerGroupId}/planner/plans?$select=id,title`) {
        return {
          value: [
            singlePlanResponseTrimmed,
            { title: singlePlanResponseTrimmed.title, id: 'RnZkyGveikyfG2u817hodJgADNwq' }
          ]
        };
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getPlanIdByTitle(planTitle, ownerGroupId);
    assert.deepStrictEqual(actual, singlePlanResponse.id);
    assert(stubMultiResults.calledOnce);
  });

  it('correctly retrieves a bucket by title', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`) {
        return multipleBucketResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getBucketByTitle(bucketName, planId);
    assert.deepStrictEqual(actual, singleBucketResponse);
  });

  it('fails to get bucket by title when it does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`) {
        return multipleBucketResponse;
      }

      return 'Invalid Request ' + opts.url;
    });

    await assert.rejects(planner.getBucketByTitle('Wrong name', planId),
      Error(`The specified bucket 'Wrong name' does not exist.`));
  });

  it('correctly returns a bucket when multiple buckets have the same title with prompt', async () => {
    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(singleBucketResponse);

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets`) {
        return {
          value: [
            singleBucketResponse,
            { ...singleBucketResponse, id: 'X5vwdkKh70-5hsFT1MJ8-pgAL4Si' }
          ]
        };
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getBucketByTitle(bucketName, planId);
    assert.deepStrictEqual(actual, singleBucketResponse);
    assert(stubMultiResults.calledOnce);
  });

  it('correctly retrieves a bucket ID by title', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets?$select=id,name`) {
        return multipleBucketResponseTrimmed;
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getBucketIdByTitle(bucketName, planId);
    assert.deepStrictEqual(actual, bucketId);
  });

  it('fails to get bucket ID by title when it does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets?$select=id,name`) {
        return multipleBucketResponseTrimmed;
      }

      return 'Invalid Request ' + opts.url;
    });

    await assert.rejects(planner.getBucketIdByTitle('Wrong name', planId),
      Error(`The specified bucket 'Wrong name' does not exist.`));
  });

  it('correctly returns a bucket ID when multiple buckets have the same title with prompt', async () => {
    const stubMultiResults = sinon.stub(cli, 'handleMultipleResultsFound').resolves(singleBucketResponse);

    const requestResult = [
      singleBucketResponse,
      { ...singleBucketResponse, id: 'X5vwdkKh70-5hsFT1MJ8-pgAL4Si' }
    ];

    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/plans/${planId}/buckets?$select=id,name`) {
        return {
          value: requestResult
        };
      }

      return 'Invalid Request ' + opts.url;
    });

    const actual = await planner.getBucketIdByTitle(bucketName, planId);
    const plansKeyValuePair = formatting.convertArrayToHashTable('id', requestResult);
    assert(stubMultiResults.calledOnceWithExactly(`Multiple buckets with name '${bucketName}' found.`, plansKeyValuePair));
    assert.deepStrictEqual(actual, bucketId);
  });
});