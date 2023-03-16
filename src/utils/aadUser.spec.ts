import * as assert from 'assert';
import * as sinon from 'sinon';
import request from '../request';
import { aadUser } from './aadUser';
import { formatting } from './formatting';
import { sinonUtil } from './sinonUtil';

const validUserName = 'john.doe@contoso.onmicrosoft.com';
const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
const userResponse = { value: [{ id: validUserId }] };
const userPrincipalNameResponse = { userPrincipalName: validUserName };

describe('utils/aadUser', () => {
  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  it('correctly get user id by upn', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      return 'Invalid Request';
    });

    const actual = await aadUser.getUserIdByUpn(validUserName);
    assert.strictEqual(actual, validUserId);
  });

  it('correctly gets a single user by Email.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(validUserName)}'&$select=id`) {
        return userResponse;
      }

      return 'Invalid Request';
    });

    const actual = await aadUser.getUserIdByEmail(validUserName);
    assert.strictEqual(actual, validUserId);
  });

  it('throws error message when no user was found with a specific upn', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return ({ value: [] });
      }

      throw 'Invalid request';
    });

    await assert.rejects(aadUser.getUserIdByUpn(validUserName), Error(`The specified user with user name ${validUserName} does not exist.`));
  });

  it('throws error message when no user was found with a specific mail', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(validUserName)}'&$select=id`) {
        return ({ value: [] });
      }

      throw `Invalid request`;
    });

    await assert.rejects(aadUser.getUserIdByEmail(validUserName), Error(`The specified user with email ${validUserName} does not exist`));
  });

  it('correctly get upn by user id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}?$select=userPrincipalName`) {
        return userPrincipalNameResponse;
      }

      return 'Invalid Request';
    });

    const actual = await aadUser.getUpnByUserId(validUserId);
    assert.strictEqual(actual, validUserName);
  });
}); 
