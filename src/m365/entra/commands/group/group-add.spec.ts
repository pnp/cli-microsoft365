import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import commands from '../../commands.js';
import aadCommands from '../../aadCommands.js';
import { Logger } from '../../../../cli/Logger.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import command from './group-add.js';
import { entraUser } from '../../../../utils/entraUser.js';
import { CommandError } from '../../../../Command.js';

describe(commands.GROUP_ADD, () => {
  const randomNumber = 0.8087050548125976;
  const userUpns = ['user1@contoso.com', 'user2@contoso.com', 'user3@contoso.com', 'user4@contoso.com', 'user5@contoso.com', 'user6@contoso.com', 'user7@contoso.com', 'user8@contoso.com', 'user9@contoso.com', 'user10@contoso.com', 'user11@contoso.com', 'user12@contoso.com', 'user13@contoso.com', 'user14@contoso.com', 'user15@contoso.com', 'user16@contoso.com', 'user17@contoso.com', 'user18@contoso.com', 'user19@contoso.com', 'user20@contoso.com', 'user21@contoso.com', 'user22@contoso.com', 'user23@contoso.com', 'user24@contoso.com', 'user25@contoso.com'];
  const userIds = ['3f2504e0-4f89-11d3-9a0c-0305e82c3301', '6dcd4ce0-4f89-11d3-9a0c-0305e82c3302', '9b76f130-4f89-11d3-9a0c-0305e82c3303', 'c835f5e0-4f89-11d3-9a0c-0305e82c3304', 'f4f3fa90-4f89-11d3-9a0c-0305e82c3305', '2230f6a0-4f8a-11d3-9a0c-0305e82c3306', '4f6df5b0-4f8a-11d3-9a0c-0305e82c3307', '7caaf4c0-4f8a-11d3-9a0c-0305e82c3308', 'a9e8f3d0-4f8a-11d3-9a0c-0305e82c3309', 'd726f2e0-4f8a-11d3-9a0c-0305e82c330a', '0484f1f0-4f8b-11d3-9a0c-0305e82c330b', '31e2f100-4f8b-11d3-9a0c-0305e82c330c', '5f40f010-4f8b-11d3-9a0c-0305e82c330d', '8c9eef20-4f8b-11d3-9a0c-0305e82c330e', 'b9fce030-4f8b-11d3-9a0c-0305e82c330f', 'e73cdf40-4f8b-11d3-9a0c-0305e82c3310', '1470ce50-4f8c-11d3-9a0c-0305e82c3311', '41a3cd60-4f8c-11d3-9a0c-0305e82c3312', '6ed6cc70-4f8c-11d3-9a0c-0305e82c3313', '9c09cb80-4f8c-11d3-9a0c-0305e82c3314', 'c93cca90-4f8c-11d3-9a0c-0305e82c3315', 'f66cc9a0-4f8c-11d3-9a0c-0305e82c3316', '2368c8b0-4f8d-11d3-9a0c-0305e82c3317', '5064c7c0-4f8d-11d3-9a0c-0305e82c3318', '7d60c6d0-4f8d-11d3-9a0c-0305e82c3319'];
  const microsoft365Group = {
    "id": "7167b488-1ffb-43f1-9547-35969469bada",
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2024-01-09T08:15:16Z",
    "creationOptions": [],
    "description": "Microsoft 365 group",
    "displayName": "Microsoft 365 Group",
    "expirationDateTime": null,
    "groupTypes": [
      "Unified"
    ],
    "isAssignableToRole": null,
    "mail": "Microsoft365Group@4wrvkx.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "Microsoft365Group",
    "membershipRule": null,
    "membershipRuleProcessingState": null,
    "onPremisesDomainName": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesNetBiosName": null,
    "onPremisesSamAccountName": null,
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "preferredLanguage": null,
    "proxyAddresses": [
      "SMTP:Microsoft365Group@4wrvkx.onmicrosoft.com"
    ],
    "renewedDateTime": "2024-01-09T08:15:16Z",
    "resourceBehaviorOptions": [],
    "resourceProvisioningOptions": [],
    "securityEnabled": true,
    "securityIdentifier": "S-1-12-1-1902621832-1139875835-2520074133-3669649812",
    "theme": null,
    "visibility": "Public",
    "onPremisesProvisioningErrors": [],
    "serviceProvisioningErrors": []
  };
  const securityGroup = {
    "id": "bc91082e-73ad-4a97-9852-e66004c7b0b6",
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2024-01-09T08:16:28Z",
    "creationOptions": [],
    "description": "Security group",
    "displayName": "Security Group",
    "expirationDateTime": null,
    "groupTypes": [],
    "isAssignableToRole": null,
    "mail": null,
    "mailEnabled": false,
    "mailNickname": "SecurityGroup",
    "membershipRule": null,
    "membershipRuleProcessingState": null,
    "onPremisesDomainName": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesNetBiosName": null,
    "onPremisesSamAccountName": null,
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "preferredLanguage": null,
    "proxyAddresses": [],
    "renewedDateTime": "2024-01-09T08:16:28Z",
    "resourceBehaviorOptions": [],
    "resourceProvisioningOptions": [],
    "securityEnabled": true,
    "securityIdentifier": "S-1-12-1-3163621422-1251439533-1625707160-3065038596",
    "theme": null,
    "visibility": "Public",
    "onPremisesProvisioningErrors": [],
    "serviceProvisioningErrors": []
  };
  const groupWithGeneratedMailNickname = {
    "id": "7167b488-1ffb-43f1-9547-35969469bada",
    "deletedDateTime": null,
    "classification": null,
    "createdDateTime": "2024-01-09T08:15:16Z",
    "creationOptions": [],
    "description": "Microsoft 365 group",
    "displayName": "Microsoft 365 Group",
    "expirationDateTime": null,
    "groupTypes": [
      "Unified"
    ],
    "isAssignableToRole": null,
    "mail": "Group808705@4wrvkx.onmicrosoft.com",
    "mailEnabled": true,
    "mailNickname": "Group808705",
    "membershipRule": null,
    "membershipRuleProcessingState": null,
    "onPremisesDomainName": null,
    "onPremisesLastSyncDateTime": null,
    "onPremisesNetBiosName": null,
    "onPremisesSamAccountName": null,
    "onPremisesSecurityIdentifier": null,
    "onPremisesSyncEnabled": null,
    "preferredDataLocation": null,
    "preferredLanguage": null,
    "proxyAddresses": [
      "SMTP:Group808705@4wrvkx.onmicrosoft.com"
    ],
    "renewedDateTime": "2024-01-09T08:15:16Z",
    "resourceBehaviorOptions": [],
    "resourceProvisioningOptions": [],
    "securityEnabled": true,
    "securityIdentifier": "S-1-12-1-1902621832-1139875835-2520074133-3669649812",
    "theme": null,
    "visibility": "Public",
    "onPremisesProvisioningErrors": [],
    "serviceProvisioningErrors": []
  };
  const addOwnersRequest = [
    {
      id: 1,
      method: 'PATCH',
      url: `/groups/${microsoft365Group.id}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'owners@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    },
    {
      id: 21,
      method: 'PATCH',
      url: `/groups/${microsoft365Group.id}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'owners@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    }
  ];
  const addMembersRequest = [
    {
      id: 1,
      method: 'PATCH',
      url: `/groups/${microsoft365Group.id}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'members@odata.bind': userIds.slice(0, 20).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    },
    {
      id: 21,
      method: 'PATCH',
      url: `/groups/${microsoft365Group.id}`,
      headers: { 'content-type': 'application/json;odata.metadata=none' },
      body: {
        'members@odata.bind': userIds.slice(20, 40).map(u => `https://graph.microsoft.com/v1.0/directoryObjects/${u}`)
      }
    }
  ];

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      entraUser.getUserIdsByUpns
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.GROUP_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notStrictEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.deepStrictEqual(alias, [aadCommands.GROUP_ADD]);
  });

  it('fails validation if the length of displayName is more than 256 characters', async () => {
    const displayName = 'lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum lorem ipsum';
    const actual = await command.validate({ options: { displayName: displayName, type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the length of mailNickname is more than 64 characters', async () => {
    const mailNickname = 'loremipsumloremipsumloremipsumloremipsumloremipsumloremipsumloremipsumlorem';
    const actual = await command.validate({ options: { displayName: 'Cli group', mailNickname: mailNickname, type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if mailNickname is not valid', async () => {
    const mailNickname = 'lorem ipsum';
    const actual = await command.validate({ options: { displayName: 'Cli group', mailNickname: mailNickname, type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerIds contains invalid GUID', async () => {
    const ownerIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerIds: ownerIds.join(','), type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if ownerUserNames contains invalid user principal name', async () => {
    const ownerUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerUserNames: ownerUserNames.join(','), type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberIds contains invalid GUID', async () => {
    const memberIds = ['7167b488-1ffb-43f1-9547-35969469bada', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', memberIds: memberIds.join(','), type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if memberUserNames contains invalid user principal name', async () => {
    const memberUserNames = ['john.doe@contoso.com', 'foo'];
    const actual = await command.validate({ options: { displayName: 'Cli group', memberUserNames: memberUserNames.join(','), type: 'security' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if visibility contains invalid value', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', visibility: 'foo', type: 'microsoft365' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if type contains invalid value', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', visibility: 'Public', type: 'foo' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if type is microsoft365 but visibility is not specified', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', type: 'microsoft365' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with ids', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerIds: userIds.join(','), memberIds: userIds.join(','), type: 'security' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when all required parameters are valid with user names', async () => {
    const actual = await command.validate({ options: { displayName: 'Cli group', ownerUserNames: userUpns.join(','), memberUserNames: userUpns.join(','), type: 'security' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('successfully creates Microsoft 365 group without owners and members', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365' } });
    assert(loggerLogSpy.calledWith(microsoft365Group));
  });

  it('successfully creates security group without owners and members', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return securityGroup;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Security Group', description: 'Security Group', mailNickname: 'SecurityGroup', type: 'security' } });
    assert(loggerLogSpy.calledWith(securityGroup));
  });

  it('successfully creates group with owners specified by ids', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', ownerIds: userIds.join(',') } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, addOwnersRequest);
  });

  it('successfully creates group with members specified by ids', async () => {
    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', memberIds: userIds.join(',') } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, addMembersRequest);
  });

  it('successfully creates group with owners specified by user names', async () => {
    sinon.stub(entraUser, 'getUserIdsByUpns').withArgs(userUpns).resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', ownerUserNames: userUpns.join(',') } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, addOwnersRequest);
  });

  it('successfully creates group with members specified by user names', async () => {
    sinon.stub(entraUser, 'getUserIdsByUpns').withArgs(userUpns).resolves(userIds);

    const postStub = sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: Array(2).fill({
            status: 204,
            body: {}
          })
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', memberUserNames: userUpns.join(','), verbose: true } });
    assert.deepStrictEqual(postStub.lastCall.args[0].data.requests, addMembersRequest);
  });

  it('successfully creates group with generated mailNickname', async () => {
    sinon.stub(Math, 'random').resolves(randomNumber);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return groupWithGeneratedMailNickname;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', visibility: 'Public', type: 'microsoft365' } });
    assert(loggerLogSpy.calledWith(groupWithGeneratedMailNickname));
  });

  it('handles API error when adding users to a group', async () => {
    sinon.stub(request, 'post').callsFake(async opts => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups') {
        return microsoft365Group;
      }

      if (opts.url === 'https://graph.microsoft.com/v1.0/$batch') {
        return {
          responses: [
            {
              id: 1,
              status: 204,
              body: {}
            },
            {
              id: 2,
              status: 400,
              body: {
                error: {
                  message: `One or more added object references already exist for the following modified properties: 'members'.`
                }
              }
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', ownerIds: userIds.join(',') } }),
      new CommandError(`One or more added object references already exist for the following modified properties: 'members'.`));
  });

  it('correctly handles API OData error', async () => {
    sinon.stub(request, 'post').rejects({
      error: {
        'odata.error': {
          code: '-1, InvalidOperationException',
          message: {
            value: 'Invalid request'
          }
        }
      }
    });

    await assert.rejects(command.action(logger, { options: { displayName: 'Microsoft 365 Group', description: 'Microsoft 365 group', mailNickname: 'Microsoft365Group', visibility: 'Public', type: 'microsoft365', ownerIds: userIds.join(',') } }),
      new CommandError('Invalid request'));
  });
});