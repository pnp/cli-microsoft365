import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./team-list');

describe(commands.TEAM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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
      appInsights.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAM_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists Microsoft Teams in the tenant', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
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
            "allowUserEditMessages": false,
            "allowUserDeleteMessages": false,
            "allowOwnerDeleteMessages": false,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "isArchived": false
      }
    ]));
  });

  it('correctly handles when listing a team a user is not member of', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315`) {
        return Promise.reject({ statusCode: 403 });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "isArchived": undefined
      }
    ]));
  });

  it('correctly handles when listing a team a user is not member of, and the graph returns an unexpected error', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Invalid request'));
  });

  it('lists Microsoft Teams in the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
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
            "allowUserEditMessages": false,
            "allowUserDeleteMessages": false,
            "allowOwnerDeleteMessages": false,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "isArchived": false
      },
      {
        "id": "13be6971-79db-4f33-9d41-b25589ca25af",
        "displayName": "Team 2",
        "description": "Team 2 description",
        "isArchived": false
      }
    ]));
  });

  it('lists joined Microsoft Teams in the tenant', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams/$entity",
          "id": id,
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
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
            "allowUserEditMessages": false,
            "allowUserDeleteMessages": false,
            "allowOwnerDeleteMessages": false,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "displayName": "Team 1",
              "description": "Team 1 description",
              "isArchived": false
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "displayName": "Team 2",
              "description": "Team 2 description",
              "isArchived": false
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { joined: true } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "isArchived": false
      },
      {
        "id": "13be6971-79db-4f33-9d41-b25589ca25af",
        "displayName": "Team 2",
        "description": "Team 2 description",
        "isArchived": false
      }
    ]));
  });

  it('lists all properties for output json', async () => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#teams/$entity",
          "id": id,
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "memberSettings": {
            "allowCreateUpdateChannels": true,
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
            "allowUserEditMessages": false,
            "allowUserDeleteMessages": false,
            "allowOwnerDeleteMessages": false,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          },
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "displayName": "Team 1",
              "description": "Team 1 description",
              "isArchived": false
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "displayName": "Team 2",
              "description": "Team 2 description",
              "isArchived": false
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith([
      {
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "isArchived": false
      },
      {
        "id": "13be6971-79db-4f33-9d41-b25589ca25af",
        "displayName": "Team 2",
        "description": "Team 2 description",
        "isArchived": false
      }
    ]));
  });
});