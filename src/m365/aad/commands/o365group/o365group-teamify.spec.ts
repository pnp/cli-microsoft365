import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./o365group-teamify');

describe(commands.O365GROUP_TEAMIFY, () => {
  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.put
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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_TEAMIFY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if both groupId and mailNickname options are not passed', (done) => {
    const actual = command.validate({
      options: {
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if both groupId and mailNickname options are passed', (done) => {
    const actual = command.validate({
      options: {
        groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        mailNickname: 'GroupName'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('validates for a correct groupId', (done) => {
    const actual = command.validate({
      options: {
        groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('fails to get o365 group when it does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq '`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('The specified Microsoft 365 Group does not exist');
    });

    command.action(logger, {
      options: {
        debug: true,
        mailNickname: 'GroupName'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified Microsoft 365 Group does not exist`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('fails when multiple groups with same name exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
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
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            },
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
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
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        mailNickname: 'GroupName'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`Multiple Microsoft 365 Groups with name GroupName found: 00000000-0000-0000-0000-000000000000,00000000-0000-0000-0000-000000000000`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Teamify o365 group by groupId', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": "8231f9f2-701f-4c6e-93ce-ecb563e3c1ee",
          "createdDateTime": null,
          "displayName": "Group Team",
          "description": "Group Team description",
          "internalId": "19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01@thread.tacv2",
          "classification": null,
          "specialization": null,
          "mailNickname": "groupname",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=8231f9f2-701f-4c6e-93ce-ecb563e3c1ee&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "isArchived": null,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": null,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowCreatePrivateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          },
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          },
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": true
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Teamify o365 group by mailNickname', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=mailNickname eq `) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#groups",
          "value": [
            {
              "@odata.id": "https://graph.microsoft.com/v2/00000000-0000-0000-0000-000000000000/directoryObjects/00000000-0000-0000-0000-000000000000/Microsoft.DirectoryServices.Group",
              "id": "00000000-0000-0000-0000-000000000000",
              "deletedDateTime": null,
              "classification": null,
              "createdDateTime": "2021-09-05T09:01:19Z",
              "creationOptions": [],
              "description": "GroupName",
              "displayName": "GroupName",
              "expirationDateTime": null,
              "groupTypes": [
                "Unified"
              ],
              "isAssignableToRole": null,
              "mail": "groupname@contoso.onmicrosoft.com",
              "mailEnabled": true,
              "mailNickname": "groupname",
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
                "SPO:SPO_00000000-0000-0000-0000-000000000000@SPO_00000000-0000-0000-0000-000000000000",
                "SMTP:groupname@contoso.onmicrosoft.com"
              ],
              "renewedDateTime": "2021-09-05T09:01:19Z",
              "resourceBehaviorOptions": [],
              "resourceProvisioningOptions": [
                "Team"
              ],
              "securityEnabled": false,
              "securityIdentifier": "S-1-12-1-71288816-1279290235-2033184675-371261341",
              "theme": null,
              "visibility": "Public",
              "onPremisesProvisioningErrors": []
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    const requestStub: sinon.SinonStub = sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": "00000000-0000-0000-0000-000000000000",
          "createdDateTime": null,
          "displayName": "Group Team",
          "description": "Group Team description",
          "internalId": "19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01@thread.tacv2",
          "classification": null,
          "specialization": null,
          "mailNickname": "groupname",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=00000000-0000-0000-0000-000000000000&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8",
          "isArchived": null,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": null,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
            "allowCreatePrivateChannels": true,
            "allowDeleteChannels": true,
            "allowAddRemoveApps": true,
            "allowCreateUpdateRemoveTabs": true,
            "allowCreateUpdateRemoveConnectors": true
          },
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          },
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": true
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, mailNickname: 'groupname' }
    }, () => {
      try {
        assert.strictEqual(requestStub.lastCall.args[0].url, 'https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/team');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee/team`) {
        return Promise.reject({
          "error": {
            "code": "NotFound",
            "message": "Failed to execute MS Graph backend request GetGroupInternalApiRequest",
            "innerError": {
              "date": "2021-06-19T03:00:13",
              "request-id": "0e3f93f6-d3f7-4d84-9eb5-dc2dda0eec0e",
              "client-request-id": "68cff2aa-b010-daa7-2467-fa8e96cbda25"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'Failed to execute MS Graph backend request GetGroupInternalApiRequest');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the groupId is not a valid GUID', () => {
    const actual = command.validate({ options: { groupId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the groupId is a valid GUID', () => {
    const actual = command.validate({ options: { groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } });
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
});