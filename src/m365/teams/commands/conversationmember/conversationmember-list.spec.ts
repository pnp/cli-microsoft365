import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
import * as os from 'os';
const command: Command = require('./conversationmember-list');

describe(commands.TEAMS_CONVERSATIONMEMBER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    Utils.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_CONVERSATIONMEMBER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if teamId and channelId are not specified', () => {
    const actual = command.validate({
      options: {
        debug: false,
      }
    });
    assert.notStrictEqual(actual, true);
  });
  
  it('fails validation if teamName and channelName are not specified', () => {
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
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
      }
    });
    assert.notStrictEqual(actual, true);
  });
  
  it('fails validatation for a incorrect channelId missing leading 19:.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '552b7125655c46d5b5b86db02ee7bfdf@thread.skype',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation for a incorrect channelId missing trailing @thread.skpye.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '19:552b7125655c46d5b5b86db02ee7bfdf@thread',
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
  
  it('validates for a correct teamId and channelId input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
        channelId: "19:eb30973b42a847a2a1df92d91e37c76a@thread.skype"
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
  
  it('validates for a correct teamId and channelName input', () => {
    const actual = command.validate({
      options: {
        teamId: "fce9e580-8bba-4638-ab5c-ab40016651e3",
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
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
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
        assert(loggerSpy.calledWith([
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

  it('lists conversation members with (teamId and channelId)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
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
        assert(loggerSpy.calledWith([
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

  it('lists conversation members with (teamName and channelName)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
      }
      
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
              {
                  "id": "47d6625d-a540-4b59-a4ab-19b787e40593",
                  "createdDateTime": null,
                  "displayName": "Human Resources",
                  "description": "Human Resources",
                  "internalId": null,
                  "classification": null,
                  "specialization": null,
                  "visibility": null,
                  "webUrl": null,
                  "isArchived": false,
                  "isMembershipLimitedToOwners": null,
                  "memberSettings": null,
                  "guestSettings": null,
                  "messagingSettings": null,
                  "funSettings": null,
                  "discoverySettings": null
              }
          ]
      });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels?$filter=displayName eq '${encodeURIComponent('Private Channel')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
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
        });
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
        assert(loggerSpy.calledWith([
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

  it('lists conversation members with (teamId and channelName)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
      }
      
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
              {
                  "id": "47d6625d-a540-4b59-a4ab-19b787e40593",
                  "createdDateTime": null,
                  "displayName": "Human Resources",
                  "description": "Human Resources",
                  "internalId": null,
                  "classification": null,
                  "specialization": null,
                  "visibility": null,
                  "webUrl": null,
                  "isArchived": false,
                  "isMembershipLimitedToOwners": null,
                  "memberSettings": null,
                  "guestSettings": null,
                  "messagingSettings": null,
                  "funSettings": null,
                  "discoverySettings": null
              }
          ]
      });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels?$filter=displayName eq '${encodeURIComponent('Private Channel')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels",
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
        });
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
        assert(loggerSpy.calledWith([
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

  
  it('lists conversation members with (teamName and channelId)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
      }
      
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
              {
                  "id": "47d6625d-a540-4b59-a4ab-19b787e40593",
                  "createdDateTime": null,
                  "displayName": "Human Resources",
                  "description": "Human Resources",
                  "internalId": null,
                  "classification": null,
                  "specialization": null,
                  "visibility": null,
                  "webUrl": null,
                  "isArchived": false,
                  "isMembershipLimitedToOwners": null,
                  "memberSettings": null,
                  "guestSettings": null,
                  "messagingSettings": null,
                  "funSettings": null,
                  "discoverySettings": null
              }
          ]
        });
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
        assert(loggerSpy.calledWith([
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
      if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams?$filter=displayName eq '${encodeURIComponent('Human Resources')}'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
              {
                  "id": "47d6625d-a540-4b59-a4ab-19b787e40593",
                  "createdDateTime": null,
                  "displayName": "Human Resources",
                  "description": "Human Resources",
                  "internalId": null,
                  "classification": null,
                  "specialization": null,
                  "visibility": null,
                  "webUrl": null,
                  "isArchived": false,
                  "isMembershipLimitedToOwners": null,
                  "memberSettings": null,
                  "guestSettings": null,
                  "messagingSettings": null,
                  "funSettings": null,
                  "discoverySettings": null
              },
              {
                "id": "47d6625d-a540-4b59-a4ab-19b787e40594",
                "createdDateTime": null,
                "displayName": "Human Resources",
                "description": "Human Resources",
                "internalId": null,
                "classification": null,
                "specialization": null,
                "visibility": null,
                "webUrl": null,
                "isArchived": false,
                "isMembershipLimitedToOwners": null,
                "memberSettings": null,
                "guestSettings": null,
                "messagingSettings": null,
                "funSettings": null,
                "discoverySettings": null
            }
          ]
        });
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
            '- 47d6625d-a540-4b59-a4ab-19b787e40594'].join(os.EOL)}`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists conversation members (json)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/47d6625d-a540-4b59-a4ab-19b787e40593/channels/19:586a8b9e36c4479bbbd378e439a96df2@thread.skype/members`) {
        return Promise.resolve({
            "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams('47d6625d-a540-4b59-a4ab-19b787e40593')/channels('19%3A586a8b9e36c4479bbbd378e439a96df2%40thread.skype')/members",
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
        });
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
        assert(loggerSpy.calledWith([
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
    sinon.stub(request, 'get').callsFake((opts) => {
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