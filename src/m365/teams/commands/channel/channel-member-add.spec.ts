import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './channel-member-add.js';
import { settingsNames } from '../../../../settingsNames.js';

describe(commands.CHANNEL_MEMBER_ADD, () => {
  //#region Mocked Responses 
  const singleTeamResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
    "value": [
      {
        "id": "47d6625d-a540-4b59-a4ab-19b787e40593",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2018-12-28T04:09:33Z",
        "createdByAppId": null,
        "description": "Human Resources",
        "displayName": "Human Resources",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "infoCatalogs": [],
        "isAssignableToRole": null,
        "mail": "hr@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "hr",
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
          "SPO:SPO_c562a29c-2afd-4b53-ae4d-f94f200de3ef@SPO_d544d1e7-d321-494b-870a-1beac97967a2",
          "SMTP:hr@sconsoto.onmicrosoft.com"
        ],
        "renewedDateTime": "2018-12-28T04:09:33Z",
        "resourceBehaviorOptions": [],
        "resourceProvisioningOptions": [
          "Team"
        ],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-1205232221-1264166208-3071912868-2466636935",
        "theme": null,
        "visibility": "Private",
        "writebackConfiguration": {
          "isEnabled": null,
          "onPremisesGroupType": null
        },
        "onPremisesProvisioningErrors": []
      }
    ]
  };

  const conversationMembersOwnerResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('d21d1577-83b5-4357-a09b-6d338c44fac4')/channels('19%3Aa0555558d9e842c3a8bae7d9d6734d7d%40thread.skype')/members/$entity",
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    "id": "MTk6YTA1NTU1NThkOWU4NDJjM2E4YmFlN2Q5ZDY3MzRkN2RAdGhyZWFkLnNreXBlIyNmNjYyMjQ2OS1hYTMzLTRjMDMtOTJmZi1hM2E1NDU0ZGY4NjQ=",
    "roles": [
      "owner"
    ],
    "displayName": "Admin",
    "userId": "f6622469-aa33-4c03-92ff-a3a5454df864",
    "email": "admin@contoso.com"
  };

  const conversationMembersResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('d21d1577-83b5-4357-a09b-6d338c44fac4')/channels('19%3Aa0555558d9e842c3a8bae7d9d6734d7d%40thread.skype')/members/$entity",
    "@odata.type": "#microsoft.graph.aadUserConversationMember",
    "id": "MTk6YTA1NTU1NThkOWU4NDJjM2E4YmFlN2Q5ZDY3MzRkN2RAdGhyZWFkLnNreXBlIyNmNjYyMjQ2OS1hYTMzLTRjMDMtOTJmZi1hM2E1NDU0ZGY4NjQ=",
    "roles": [
      "owner"
    ],
    "displayName": "Admin",
    "userId": "f6622469-aa33-4c03-92ff-a3a5454df864",
    "email": "admin@contoso.com"
  };

  const singleChannelResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
    "@odata.count": 1,
    "value": [
      {
        "id": "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        "displayName": "Private Channel",
        "description": null,
        "isFavoriteByDefault": null,
        "email": "",
        "webUrl": "https://teams.microsoft.com/l/channel/19%3a586a8b9e36c4479bbbd378e439a96df2%40thread.skype/Private+Channel?groupId=47d6625d-a540-4b59-a4ab-19b787e40593&tenantId=d544d1e7-d321-494b-870a-1beac97967a2",
        "membershipType": "private"
      }
    ]
  };

  const channelIdResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels/$entity",
    "id": "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
    "displayName": "Private Channel",
    "description": null,
    "isFavoriteByDefault": null,
    "email": "",
    "webUrl": "https://teams.microsoft.com/l/channel/19%3a586a8b9e36c4479bbbd378e439a96df2%40thread.skype/Private+Channel?groupId=47d6625d-a540-4b59-a4ab-19b787e40593&tenantId=d544d1e7-d321-494b-870a-1beac97967a2",
    "membershipType": "private"
  };

  const singleUserResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users",
    "value": [
      {
        "businessPhones": [],
        "displayName": "Admin",
        "givenName": "Admin",
        "jobTitle": "Software Developer",
        "mail": "admin@contoso.com",
        "mobilePhone": null,
        "officeLocation": null,
        "preferredLanguage": null,
        "surname": "Admin",
        "userPrincipalName": "admin@contoso.com",
        "id": "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    ]
  };

  const multipleUserResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users",
    "value": [
      {
        "businessPhones": [
          "4250000000"
        ],
        "displayName": "Admin",
        "givenName": "Admin",
        "jobTitle": "SharePoint Consultant",
        "mail": "admin@contoso.com",
        "mobilePhone": null,
        "officeLocation": null,
        "preferredLanguage": "en-US",
        "surname": "Admin",
        "userPrincipalName": "admin@contoso.com",
        "id": "4cb2b035-ad76-406c-bdc4-6c72ad403a22"
      },
      {
        "businessPhones": [],
        "displayName": "Admin",
        "givenName": "Admin",
        "jobTitle": null,
        "mail": "admin2@contoso.com",
        "mobilePhone": null,
        "officeLocation": null,
        "preferredLanguage": null,
        "surname": "Admin",
        "userPrincipalName": "admin2@contoso.com",
        "id": "662c9a98-1e96-44d2-b5ef-4933004200f8"
      }
    ]
  };

  const noUserResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users",
    "value": []
  };
  //#endregion

  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
    sinon.stub(Cli.getInstance(), 'getSettingWithDefaultValue').callsFake((settingName: string, defaultValue: any) => {
      if (settingName === 'prompt') {
        return false;
      }

      return defaultValue;
    });
  });

  beforeEach(() => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return conversationMembersOwnerResponse;
      }

      throw 'Invalid request';
    });
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return channelIdResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('Human Resources')}'`) {
        return singleTeamResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter('Admin')}'`) {
        return singleUserResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${formatting.encodeQueryParameter('Private Channel')}'`) {
        return singleChannelResponse;
      }

      throw 'Invalid request';
    });
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
    loggerLogSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      cli.getSettingWithDefaultValue,
      Cli.handleMultipleResultsFound
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CHANNEL_MEMBER_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect channelId missing leading 19:.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '586a8b9e36c4479bbbd378e439a96df2@thread.skype',
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skype.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:586a8b9e36c4479bbbd378e439a96df2',
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct teamId, channelId, and userIds input', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userIds input', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamName, channelName, and userIds input', async () => {
    const actual = await command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userIds: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelId, and userDisplayNames input', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayNames: "admin.contoso.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userDisplayNames input', async () => {
    const actual = await command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userDisplayNames: "admin.contoso.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamName, channelName, and userDisplayNames input', async () => {
    const actual = await command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userDisplayNames: "admin.contoso.com"
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('adds conversation members using teamName, channelId, and userIds', async () => {
    await command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userIds: "admin@contoso.com",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('adds conversation members using teamId, channelName, and userIds', async () => {
    await command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Private Channel",
        userIds: "admin@contoso.com",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('adds conversation members using teamName, channelName, and userIds', async () => {
    await command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userIds: "admin@contoso.com",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('adds conversation members using teamName, channelId, and userDisplayNames', async () => {
    await command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayNames: "Admin",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('adds conversation members using teamId, channelName, and userDisplayNames', async () => {
    await command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Private Channel",
        userDisplayNames: "Admin",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('adds conversation members using teamName, channelName, and userDisplayNames', async () => {
    sinonUtil.restore(request.post);
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return conversationMembersResponse;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userDisplayNames: "Admin"
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('fails adding conversation members with invalid channelName', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('Human Resources')}'`) {
        return singleTeamResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${formatting.encodeQueryParameter('Other Private Channel')}'`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
          "@odata.count": 0,
          "value": []
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Other Private Channel"
      }
    } as any), new CommandError(`The specified channel 'Other Private Channel' does not exist in the Microsoft Teams team with ID '47d6625d-a540-4b59-a4ab-19b787e40593'`));
  });

  it('fails to get channel when channel does is not private', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${formatting.encodeQueryParameter('Human Resources')}'`) {
        return singleTeamResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${formatting.encodeQueryParameter('Other Channel')}'`) {
        return {
          "value": [
            {
              "name": "Other Channel",
              "membershipType": "standard"
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Other Channel"
      }
    } as any), new CommandError('The specified channel is not a private channel'));
  });

  it('fails when group has no team', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": []
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamName: 'Team Name',
        channelName: "Other Channel"
      }
    } as any), new CommandError('The specified team does not exist in the Microsoft Teams'));
  });

  it('fails adding conversation members with multiple userDisplayNames', async () => {
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((settingName, defaultValue) => {
      if (settingName === settingsNames.prompt) {
        return false;
      }

      return defaultValue;
    });

    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter('Admin')}'`) {
        return multipleUserResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return channelIdResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayNames: "Admin"
      }
    } as any), new CommandError("Multiple users with display name 'Admin' found. Found: 4cb2b035-ad76-406c-bdc4-6c72ad403a22, 662c9a98-1e96-44d2-b5ef-4933004200f8."));
  });

  it('handles selecting single result when multiple userDisplayNames with the specified name found and cli is set to prompt', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter('Admin')}'`) {
        return multipleUserResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return channelIdResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(Cli, 'handleMultipleResultsFound').resolves(singleUserResponse);

    await command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayNames: "Admin",
        owner: true
      }
    });
    assert(loggerLogSpy.notCalled);
  });

  it('fails adding conversation members when no users are found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${formatting.encodeQueryParameter('Admin')}'`) {
        return noUserResponse;
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${formatting.encodeQueryParameter('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${formatting.encodeQueryParameter('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return channelIdResponse;
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayNames: "Admin"
      }
    } as any), new CommandError("The specified user 'Admin' does not exist"));
  });

  it('correctly handles error when adding conversation members', async () => {
    sinonUtil.restore(request.get);
    const error = {
      "error": {
        "code": "UnknownError",
        "message": "An error has occurred",
        "innerError": {
          "date": "2022-02-14T13:27:37",
          "request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c",
          "client-request-id": "77e0ed26-8b57-48d6-a502-aca6211d6e7c"
        }
      }
    };
    sinon.stub(request, 'get').rejects(error);

    await assert.rejects(command.action(logger, {
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        userDisplayNames: "Admin"
      }
    } as any), new CommandError('An error has occurred'));
  });
});
