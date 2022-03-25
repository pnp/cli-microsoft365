import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./conversationmember-list');

describe(commands.CONVERSATIONMEMBER_LIST, () => {
  //#region Mocked Responses 
  const multipleTeamsResponse: any = {
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
  };

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

  const conversationMembersResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
    "@odata.count": 2,
    "value": [
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
        "roles": [
          "owner"
        ],
        "displayName": "Mary Thompson",
        "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
        "email": "mary@contoso.com"
      },
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
        "roles": [],
        "displayName": "John Smith",
        "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
        "email": "john@contoso.com"
      }
    ]
  };

  const singleChannelResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
    "@odata.count": 1,
    "value": [
      {
        "id": "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
        "createdDateTime": null,
        "displayName": "Private Channel",
        "description": null,
        "isFavoriteByDefault": null,
        "email": "",
        "webUrl": "https://teams.microsoft.com/l/channel/19%3a586a8b9e36c4479bbbd378e439a96df2%40thread.skype/Private+Channel?groupId=47d6625d-a540-4b59-a4ab-19b787e40593&tenantId=d544d1e7-d321-494b-870a-1beac97967a2",
        "membershipType": "private",
        "moderationSettings": null
      }
    ]
  };

  const channelIdResponse: any = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels/$entity",
    "id": "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype",
    "displayName": "Private Channel",
    "description": null,
    "email": "",
    "webUrl": "https://teams.microsoft.com/l/channel/19%3a586a8b9e36c4479bbbd378e439a96df2%40thread.skype/Private+Channel?groupId=47d6625d-a540-4b59-a4ab-19b787e40593&tenantId=d544d1e7-d321-494b-870a-1beac97967a2",
    "membershipType": "private"
  };

  const channelIdErrorResponse: any = {
    "error": {
      "code": "NotFound",
      "message": "Failed to execute Skype backend request GetThreadS2SRequest.",
      "innerError": {
        "date": "2020-11-05T15:30:50",
        "request-id": "bf7c27d4-38d1-42a8-af93-03e5446af010",
        "client-request-id": "89f8859a-bc75-36ce-b4ca-035a6889844d"
      }
    }
  };
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
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CONVERSATIONMEMBER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId and channelId are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if teamName and channelName are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validatation for a incorrect channelId missing leading 19:.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '586a8b9e36c4479bbbd378e439a96df2@thread.skype'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:586a8b9e36c4479bbbd378e439a96df2'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the channelName is empty', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: ""
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the teamName is empty', () => {
    const actual = command.validate({
      options: {
        teamName: "",
        channelName: "Private Channel"
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('fails validation if teamName and teamId are specified', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        teamName: "Human Resources",
        channelName: "Private Channel"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if channelName and channelId are specified', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct teamId and channelId input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamId and channelName input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelName: "Private Channel"
      }
    });
    assert.strictEqual(actual, true);
  });

  it('validates for a correct teamName and channelName input', () => {
    const actual = command.validate({
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel"
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

  it('lists conversation members (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members with teamId and channelId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members with teamName and channelName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${encodeURIComponent('Private Channel')}'`) {
        return Promise.resolve(singleChannelResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelName: "Private Channel"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members with teamId and channelName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels?$filter=displayName eq '${encodeURIComponent('Private Channel')}'`) {
        return Promise.resolve(singleChannelResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelName: "Private Channel"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails listing conversation members with invalid teamName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Other Human Resources')}'`) {
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
          JSON.stringify(new CommandError(`The specified team 'Other Human Resources' does not exist in the Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails listing conversation members with invalid channelName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
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

  it('fails listing conversation members with invalid channelId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:whatever@thread.skype')}`) {
        return Promise.reject(channelIdErrorResponse);
      }

      return Promise.reject('Invalid Request 123');
    });

    command.action(logger, {
      options: {
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:whatever@thread.skype"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(
          JSON.stringify(err),
          JSON.stringify(new CommandError(`The specified channel '19:whatever@thread.skype' does not exist or is invalid in the Microsoft Teams team with ID '47d6625d-a540-4b59-a4ab-19b787e40593'`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members with teamName and channelId', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve(singleTeamResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/${encodeURIComponent('47d6625d-a540-4b59-a4ab-19b787e40593')}/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        teamName: "Human Resources",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members with multiple teamName', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
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
          JSON.stringify(new CommandError(`Multiple Microsoft Teams teams with name 'Human Resources' found: 47d6625d-a540-4b59-a4ab-19b787e40593,5b1fac18-4ae3-43b4-9ca8-e27c7f44b65f`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members (json)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}/members`) {
        return Promise.resolve(conversationMembersResponse);
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/${encodeURIComponent('19:586a8b9e36c4479bbbd378e439a96df2@thread.skype')}`) {
        return Promise.resolve(channelIdResponse);
      }

      return Promise.reject('Invalid Request');
    });

    command.action(logger, {
      options: {
        output: "json",
        teamId: "47d6625d-a540-4b59-a4ab-19b787e40593",
        channelId: "19:586a8b9e36c4479bbbd378e439a96df2@thread.skype"
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith([
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyM1YzcwNTI4OC1lZDdmLTQ0ZmMtYWYwYS1hYzE2NDQxOTkwMWM=",
            "roles": [
              "owner"
            ],
            "displayName": "Mary Thompson",
            "userId": "5c705288-ed7f-44fc-af0a-ac164419901c",
            "email": "mary@contoso.com"
          },
          {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "id": "MTk6NTg2YThiOWUzNmM0NDc5YmJiZDM3OGU0MzlhOTZkZjJAdGhyZWFkLnNreXBlIyMzZjhkYjI5NC03ODE0LTQzZTYtOGE1NC1hZDUxM2YzYTA2ZTE=",
            "roles": [],
            "displayName": "John Smith",
            "userId": "3f8db294-7814-43e6-8a54-ad513f3a06e1",
            "email": "john@contoso.com"
          }
        ]));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing conversation members', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
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