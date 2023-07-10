import assert from 'assert';
import sinon from 'sinon';
import request from '../request.js';
import { entraUser } from './entraUser.js';
import { formatting } from './formatting.js';
import { sinonUtil } from './sinonUtil.js';
import { Logger } from '../cli/Logger.js';

const validUserName = 'john.doe@contoso.onmicrosoft.com';
const validUserId = '2056d2f6-3257-4253-8cfc-b73393e414e5';
const userResponse = { value: [{ id: validUserId }] };
const userPrincipalNameResponse = { userPrincipalName: validUserName };

describe('utils/entraUser', () => {
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
      request.post
    ]);
  });

  it('correctly get user id by upn', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return userResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraUser.getUserIdByUpn(validUserName);
    assert.strictEqual(actual, validUserId);
  });

  it('correctly gets a single user by Email.', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(validUserName)}'&$select=id`) {
        return userResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraUser.getUserIdByEmail(validUserName);
    assert.strictEqual(actual, validUserId);
  });

  it('throws error message when no user was found with a specific upn', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${formatting.encodeQueryParameter(validUserName)}'&$select=Id`) {
        return ({ value: [] });
      }

      throw 'Invalid request';
    });

    await assert.rejects(entraUser.getUserIdByUpn(validUserName), Error(`The specified user with user name ${validUserName} does not exist.`));
  });

  it('correctly gets user ids by upns', async () => {
    const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
    const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

    let batch = -1;
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: userIds.slice(++batch * 20, batch * 20 + 20).map(userId => ({
            status: 200,
            body: {
              id: userId
            }
          }))
        };
      }

      throw 'Invalid request';
    });

    const actual = await entraUser.getUserIdsByUpns(userUpns);
    assert.deepStrictEqual(postStub.firstCall.args[0].data.requests, userUpns.slice(0, 20).map((upn, i) => ({ id: i + 1, method: 'GET', url: `/users/${formatting.encodeQueryParameter(upn)}?$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, userUpns.slice(20, 40).map((upn, i) => ({ id: i + 1, method: 'GET', url: `/users/${formatting.encodeQueryParameter(upn)}?$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(actual, userIds);
  });

  it('correctly throws error when no user was found with a specific upn', async () => {
    const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
    const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

    let counter = 0;
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: userIds.slice(counter, counter + 20).map(userId => {
            if (counter++ < userUpns.length - 1) {
              return {
                status: 200,
                body: {
                  id: userId
                }
              };
            }
            else {
              return {
                id: counter % 20,
                status: 404,
                body: {
                  error: {
                    message: 'Resource does not exist or one of its queried reference-property objects are not present.'
                  }
                }
              };
            }
          })
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(entraUser.getUserIdsByUpns(userUpns), Error(`The specified user with user name '${userUpns[userUpns.length - 1]}' does not exist.`));
  });

  it('throws error message when no user was found with a specific mail', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=mail eq '${formatting.encodeQueryParameter(validUserName)}'&$select=id`) {
        return ({ value: [] });
      }

      throw `Invalid request`;
    });

    await assert.rejects(entraUser.getUserIdByEmail(validUserName), Error(`The specified user with email ${validUserName} does not exist`));
  });

  it('correctly get upn by user id', async () => {
    sinon.stub(request, 'get').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/${validUserId}?$select=userPrincipalName`) {
        return userPrincipalNameResponse;
      }

      return 'Invalid Request';
    });

    const actual = await entraUser.getUpnByUserId(validUserId, logger, true);
    assert.strictEqual(actual, validUserName);
  });

  it('correctly gets user ids by mails', async () => {
    const userMails = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
    const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

    let batch = -1;
    const postStub = sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: userIds.slice(++batch * 20, batch * 20 + 20).map(userId => ({
            status: 200,
            body: {
              id: userId
            }
          }))
        };
      }

      throw 'Invalid request';
    });

    const actual = await entraUser.getUserIdsByEmails(userMails);
    assert.deepStrictEqual(postStub.firstCall.args[0].data.requests, userMails.slice(0, 20).map((mail, i) => ({ id: i + 1, method: 'GET', url: `/users?$filter=mail eq '${formatting.encodeQueryParameter(mail)}'&$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, userMails.slice(20, 40).map((mail, i) => ({ id: i + 1, method: 'GET', url: `/users?$filter=mail eq '${formatting.encodeQueryParameter(mail)}'&$select=id`, headers: { accept: 'application/json;odata.metadata=none' } })));
    assert.deepStrictEqual(actual, userIds);
  });

  it('correctly throws error when no user was found with a specific mail', async () => {
    const userEmails = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
    const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];

    let counter = 0;
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === `https://graph.microsoft.com/v1.0/$batch`) {
        return {
          responses: userIds.slice(counter, counter + 20).map(userId => {
            if (counter++ < userEmails.length - 1) {
              return {
                status: 200,
                body: {
                  id: userId
                }
              };
            }
            else {
              return {
                id: counter % 20,
                status: 404,
                body: {
                  error: {
                    message: 'Resource does not exist or one of its queried reference-property objects are not present.'
                  }
                }
              };
            }
          })
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(entraUser.getUserIdsByEmails(userEmails), Error(`The specified user with mail '${userEmails[userEmails.length - 1]}' does not exist.`));
  });
}); 
