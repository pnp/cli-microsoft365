import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./channel-membership-list');

describe(commands.CHANNEL_MEMBERSHIP_LIST, () => {
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
    (command as any).items = [];
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
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CHANNEL_MEMBERSHIP_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both teamId and teamName options are not passed', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both teamId and teamName options are passed', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        teamName: 'Team Name'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both channelId and channelName options are not passed', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both channelId and channelName options are passed', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        channelName: 'Channel Name'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'roles', 'displayName', 'userId', 'email']);
  });

  it('fails validation when invalid role specified', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'Invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, channelId and Owner role specified', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'owner'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, channelId and Member role specified', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'member'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, channelId and Guest role specified', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'guest'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails to get team when team does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified team does not exist in the Microsoft Teams');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified team does not exist in the Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when group has no team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2020-10-11T09:35:26Z",
              "creationOptions": [
                "ExchangeProvisioningFlags:3552"
              ],
              "description": "Team Description",
              "displayName": "Team Name",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "TeamName@contoso.com",
              "mailEnabled": true,
              "mailNickname": "TeamName",
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
                "SPO:SPO_97df7113-c3f3-447f-8010-9f88eb0fc7f1@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:TeamName@contoso.com"
              ],
              "renewedDateTime": "2020-10-11T09:35:26Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-1927732186-1159088485-2915259540-28248825",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }

      return Promise.reject('The specified team does not exist in the Microsoft Teams');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified team does not exist in the Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when multiple teams with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2020-10-11T09:35:26Z",
              "creationOptions": [
                "Team",
                "ExchangeProvisioningFlags:3552"
              ],
              "description": "Team Description",
              "displayName": "Team Name",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "TeamName@contoso.com",
              "mailEnabled": true,
              "mailNickname": "TeamName",
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
                "SPO:SPO_97df7113-c3f3-447f-8010-9f88eb0fc7f1@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:TeamName@contoso.com"
              ],
              "renewedDateTime": "2020-10-11T09:35:26Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-1927732186-1159088485-2915259540-28248825",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            },
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:14:38Z",
              "creationOptions": [],
              "description": "Team Description",
              "displayName": "Team Name",
              "expirationDateTime": null,
              "groupTypes": [],
              "isAssignableToRole": null,
              "mail": null,
              "mailEnabled": false,
              "mailNickname": "00000000-0000-0000-0000-000000000000",
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
              "renewedDateTime": "2021-09-05T09:14:38Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": true,
              "securityIdentifier": "S-1-12-1-4278539468-1089637032-1626171811-2046493509",
              "theme": null,
              "visibility": null,
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }

      return Promise.reject('Multiple Microsoft Teams teams with name Team Name found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamName: 'Team Name'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Microsoft Teams teams with name Team Name found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get teams id by team name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2020-10-11T09:35:26Z",
              "creationOptions": [
                "Team",
                "ExchangeProvisioningFlags:3552"
              ],
              "description": "Team Description",
              "displayName": "Team Name",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "TeamName@contoso.com",
              "mailEnabled": true,
              "mailNickname": "TeamName",
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
                "SPO:SPO_97df7113-c3f3-447f-8010-9f88eb0fc7f1@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:TeamName@contoso.com"
              ],
              "renewedDateTime": "2020-10-11T09:35:26Z",
              "resourceBehaviorOptions": [
                "HideGroupInOutlook",
                "SubscribeMembersToCalendarEventsDisabled",
                "WelcomeEmailDisabled"
              ],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-1927732186-1159088485-2915259540-28248825",
              "theme": null,
              "visibility": "Private",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members') > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamName: 'Team name',
        channelId: '00:00000000000000000000000000000000@thread.skype'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          []
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly get channel id by channel name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "00:00000000000000000000000000000000@thread.skype",
              "createdDateTime": "2000-01-01T00:00:00.000Z",
              "displayName": "General",
              "description": "Test Team",
              "isFavoriteByDefault": null,
              "email": "00000000.tenant.onmicrosoft.com@emea.teams.ms",
              "webUrl": "https://teams.microsoft.com/l/channel/00:00000000000000000000000000000000@thread.skype/General?groupId=00000000-0000-0000-0000-000000000000&tenantId=00000000-0000-0000-0000-000000000001",
              "membershipType": "standard"
            }
          ]
        });
      }

      if ((opts.url as string).indexOf('/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members') > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Channel Name'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          []
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails to get channel when channel does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/teams/00000000-0000-0000-0000-000000000000/channels?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": []
        });
      }

      return Promise.reject('The specified channel does not exist in the Microsoft Teams team');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: "Channel name"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving all teams', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000'
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

  it('outputs all data in json output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when filtering on member role is incorrect', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'member'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when filtering on owner role is incorrect', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'owner'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when filtering on guest role is incorrect', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/00:00000000000000000000000000000000@thread.skype/members`) {
        return Promise.resolve({
          value: [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzkxZmYyZTE3LTg0ZGUtNDU1YS04ODE1LTUyYjIxNjgzZjY0ZQ==",
              "roles": [],
              "displayName": "User 1",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000001",
              "email": "user1@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjI2IyMDkxZTE4LTc4ODItNGVmZS1iN2QxLTkwNzAzZjVhNWM2NQ==",
              "roles": [
                "owner"
              ],
              "displayName": "User 2",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000002",
              "email": "user2@domainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            },
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000',
        channelId: '00:00000000000000000000000000000000@thread.skype',
        role: 'guest'
      }
    }, () => {
      try {
        assert(loggerLogSpy.calledWith(
          [
            {
              "id": "MCMjMiMjZjU4NTk3NTgtMzE3YS00NTMzLTg3MDgtNDU3ODFlOTgzYzZhIyMxOTpkNTdmY2ZmNGMzMjE0MDVhYjY5YzJhZWVlMTIzODllMkB0aHJlYWQuc2t5cGUjIzg0OTg3NjNmLTJjYTItNGRmNy05OTBhLWZkNjg4NTJkOTVmOA==",
              "roles": [
                "guest"
              ],
              "displayName": "User 3",
              "visibleHistoryStartDateTime": "0001-01-01T00:00:00Z",
              "userId": "00000000-0000-0000-0000-000000000003",
              "email": "user3@externaldomainname.com",
              "tenantId": "00000000-0000-0000-0000-000000000000"
            }
          ]
        ));

        done();
      }
      catch (e) {
        done(e);
      }
    });
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
});
