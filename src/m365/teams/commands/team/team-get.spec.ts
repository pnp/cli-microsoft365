import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./team-get');

describe(commands.TEAM_GET, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TEAM_GET), true);
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
        id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844',
        name: 'Team Name'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = command.validate({ options: { id: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = command.validate({ options: { id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } });
    assert.strictEqual(actual, true);
  });

  it('fails to get team information due to wrong team id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/1caf7dcd-7e83-4c3a-94f7-932a1299c843`) {
        return Promise.reject({
          "error": {
            "code": "NotFound",
            "message": "No team found with Group Id 1caf7dcd-7e83-4c3a-94f7-932a1299c843",
            "innerError": {
              "message": "No team found with Group Id 1caf7dcd-7e83-4c3a-94f7-932a1299c843",
              "code": "ItemNotFound",
              "innerError": {},
              "date": "2021-09-23T01:26:41",
              "request-id": "717697d2-b63d-422f-863c-d74d0c1c8c6f"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c843' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'No team found with Group Id 1caf7dcd-7e83-4c3a-94f7-932a1299c843');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when team name does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": []
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Team Name'
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

  it('retrieves information about the specified Microsoft Team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {

        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "createdDateTime": "2017-11-29T03:27:05Z",
          "displayName": "Finance",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "classification": null,
          "specialization": "none",
          "visibility": "Public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": false
          },
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

    command.action(logger, { options: { debug: false, id: '1caf7dcd-7e83-4c3a-94f7-932a1299c844' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "createdDateTime": "2017-11-29T03:27:05Z",
          "displayName": "Finance",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "classification": null,
          "specialization": "none",
          "visibility": "Public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": false
          },
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
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves information about the specified Microsoft Teams team by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {

      if ((opts.url as string).indexOf(`/v1.0/groups?$filter=displayName eq '`) > -1) {
        return Promise.resolve({
          "value": [
            {
              "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }

      if (opts.url === `https://graph.microsoft.com/v1.0/teams/1caf7dcd-7e83-4c3a-94f7-932a1299c844`) {

        return Promise.resolve({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "createdDateTime": "2017-11-29T03:27:05Z",
          "displayName": "Finance",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "classification": null,
          "specialization": "none",
          "visibility": "Public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": false
          },
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

    command.action(logger, { options: { debug: false, name: 'Finance' } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "id": "1caf7dcd-7e83-4c3a-94f7-932a1299c844",
          "createdDateTime": "2017-11-29T03:27:05Z",
          "displayName": "Finance",
          "description": "This is the Contoso Finance Group. Please come here and check out the latest news, posts, files, and more.",
          "classification": null,
          "specialization": "none",
          "visibility": "Public",
          "webUrl": "https://teams.microsoft.com/l/team/19:ASjdflg-xKFnjueOwbm3es6HF2zx3Ki57MyfDFrjeg01%40thread.tacv2/conversations?groupId=1caf7dcd-7e83-4c3a-94f7-932a1299c844&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": false
          },
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
        }));
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
