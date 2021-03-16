import * as assert from 'assert';
import * as os from 'os';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./conversationmember-add');

describe(commands.TEAMS_CONVERSATIONMEMBER_ADD, () => {
  //#region Mocked Responses 
  const multipleTeamsResponse: any = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#groups",
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
      },
      {
        "id": "5b1fac18-4ae3-43b4-9ca8-e27c7f44b65f",
        "deletedDateTime": null,
        "classification": null,
        "createdDateTime": "2020-10-28T21:50:12Z",
        "createdByAppId": "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe",
        "description": "Human Resources",
        "displayName": "Human Resources",
        "expirationDateTime": null,
        "groupTypes": [
          "Unified"
        ],
        "infoCatalogs": [],
        "isAssignableToRole": null,
        "mail": "HumanResources@contoso.onmicrosoft.com",
        "mailEnabled": true,
        "mailNickname": "HumanResources",
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
          "SPO:SPO_4bb60dfd-0d1d-4242-9e50-cfb41c37d022@SPO_d544d1e7-d321-494b-870a-1beac97967a2",
          "SMTP:HumanResources@contoso.onmicrosoft.com"
        ],
        "renewedDateTime": "2020-10-28T21:50:12Z",
        "resourceBehaviorOptions": [
          "HideGroupInOutlook",
          "SubscribeMembersToCalendarEventsDisabled",
          "WelcomeEmailDisabled"
        ],
        "resourceProvisioningOptions": [
          "Team"
        ],
        "securityEnabled": false,
        "securityIdentifier": "S-1-12-1-1528802328-1135889123-2095229084-1605780607",
        "theme": null,
        "visibility": "Public",
        "writebackConfiguration": {
          "isEnabled": null,
          "onPremisesGroupType": null
        },
        "onPremisesProvisioningErrors": []
      }
    ]
  }

  const singleTeamResponse: any = {
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#groups",
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
  }

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
  }

  const channelIdResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels/$entity",
    "id": "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
    "displayName": "Private Channel",
    "description": null,
    "isFavoriteByDefault": null,
    "email": "",
    "webUrl": "https://teams.microsoft.com/l/channel/19%3a586a8b9e36c4479bbbd378e439a96df2%40thread.skype/Private+Channel?groupId=47d6625d-a540-4b59-a4ab-19b787e40593&tenantId=d544d1e7-d321-494b-870a-1beac97967a2",
    "membershipType": "private"
  }

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
  }

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
  }

  const noUserResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users",
    "value": []
  }
  //#endregion

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersOwnerResponse);
      }

      return Promise.reject('Invalid Request');
    });
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${encodeURIComponent('Admin')}'`) {
        return Promise.resolve(singleUserResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${encodeURIComponent('Private Channel')}'`) {
        return Promise.resolve(singleChannelResponse);
      }

      return Promise.reject('Invalid Request');
    });
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAMS_CONVERSATIONMEMBER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId, channelId, and userId are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false,
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamName, channelName, and userDisplayName are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false,
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validatation for a incorrect channelId missing leading 19:.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '586a8b9e36c4479bbbd378e439a96df2@thread.skype',
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:586a8b9e36c4479bbbd378e439a96df2',
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the channelName is empty', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamName is empty', () => {
    const actual = command.validate({
      options: {
        teamName: "",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the userDisplayName is empty', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userDisplayName: ""
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamName and teamId are specified', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        teamName: "Human Resources",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelName and channelId are specified', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if userId and userDisplayName are specified', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1",
        userDisplayName: "admin@contoso.com"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct teamId, channelId, and userId input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userId input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamName, channelName, and userId input', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userId input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userId: "f410f714-29e3-43f7-874d-d7d35c33eaf1"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelId, and userDisplayName input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayName: "admin.contoso.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userDisplayName input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userDisplayName: "admin.contoso.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamName, channelName, and userDisplayName input', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userDisplayName: "admin.contoso.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId, channelName, and userDisplayName input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel",
        userDisplayName: "admin.contoso.com"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('adds conversation members using teamName, channelId, and userId', (done) => {
    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userId: "admin@contoso.com",
        owner: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds conversation members using teamId, channelName, and userId', (done) => {
    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Private Channel",
        userId: "admin@contoso.com",
        owner: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds conversation members using teamName, channelName, and userId', (done) => {
    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userId: "admin@contoso.com",
        owner: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds conversation members using teamName, channelId, and userDisplayName', (done) => {
    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayName: "Admin",
        owner: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds conversation members using teamId, channelName, and userDisplayName', (done) => {
    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Private Channel",
        userDisplayName: "Admin",
        owner: true
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds conversation members using teamName, channelName, and userDisplayName', (done) => {
    Utils.restore(request.post);
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        userDisplayName: "Admin"
      }
    }, () => {
      try {
        assert(loggerLogSpy.notCalled);

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails adding conversation members with invalid teamName', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '${encodeURIComponent('Other Human Resources')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 0,
          "value": []
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamName: "Other Human Resources",
        channelName: "Other Private Channel"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`The specified team 'Other Human Resources' does not exist in Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails adding conversation members with invalid channelName', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${encodeURIComponent('Other Private Channel')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
          "@odata.count": 0,
          "value": []
        });
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Other Private Channel"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`The specified channel 'Other Private Channel' does not exist in the Microsoft Teams team with ID '47d6625d-a540-4b59-a4ab-19b787e40593'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails adding conversation members with multiple teamName', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team') and displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(multipleTeamsResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`Multiple Microsoft Teams with name 'Human Resources' found. Please disambiguate:${os.EOL}${[
            '- 47d6625d-a540-4b59-a4ab-19b787e40593',
            '- 5b1fac18-4ae3-43b4-9ca8-e27c7f44b65f'].join(os.EOL)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails adding conversation members with multiple userDisplayNames', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${encodeURIComponent('Admin')}'`) {
        return Promise.resolve(multipleUserResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayName: "Admin"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`Multiple users with display name 'Admin' found. Please disambiguate:${os.EOL}${[
            '- 4cb2b035-ad76-406c-bdc4-6c72ad403a22',
            '- 662c9a98-1e96-44d2-b5ef-4933004200f8'].join(os.EOL)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails adding conversation members when no users are found', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users?$filter=displayName eq '${encodeURIComponent('Admin')}'`) {
        return Promise.resolve(noUserResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        userDisplayName: "Admin"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`The specified user 'Admin' does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when adding conversation members', (done) => {
    Utils.restore(request.get);
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype",
        userDisplayName: "Admin"
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});