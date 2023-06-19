import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./team-list');

describe(commands.TEAM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TEAM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'displayName', 'isArchived', 'description']);
  });

  it('lists Microsoft Teams in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "createdDateTime": "2022-12-08T09:17:55.039Z",
          "displayName": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1" : "Team 2",
          "description": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1 description" : "Team 2 description",
          "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
          "classification": null,
          "specialization": "none",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": true
          },
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: {} });
    assert(loggerLogSpy.calledWith([
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "createdDateTime": "2022-12-08T09:17:55.039Z",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
        "classification": null,
        "specialization": "none",
        "visibility": "public",
        "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
        "isArchived": false,
        "isMembershipLimitedToOwners": false,
        "discoverySettings": {
          "showInTeamsSearchAndSuggestions": true
        },
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
      }
    ]));
  });

  it('correctly handles when listing a team a user is not member of', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315`) {
        throw { statusCode: 403 };
      }

      throw 'Invalid request';
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        };
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('Invalid request'));
  });

  it('lists Microsoft Teams in the tenant (debug)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
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
        };
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "createdDateTime": "2022-12-08T09:17:55.039Z",
          "displayName": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1" : "Team 2",
          "description": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1 description" : "Team 2 description",
          "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
          "classification": null,
          "specialization": "none",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": true
          },
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { debug: true } });
    assert(loggerLogSpy.calledWith([
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "createdDateTime": "2022-12-08T09:17:55.039Z",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
        "classification": null,
        "specialization": "none",
        "visibility": "public",
        "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
        "isArchived": false,
        "isMembershipLimitedToOwners": false,
        "discoverySettings": {
          "showInTeamsSearchAndSuggestions": true
        },
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
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
        "id": "13be6971-79db-4f33-9d41-b25589ca25af",
        "createdDateTime": "2022-12-08T09:17:55.039Z",
        "displayName": "Team 2",
        "description": "Team 2 description",
        "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
        "classification": null,
        "specialization": "none",
        "visibility": "public",
        "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
        "isArchived": false,
        "isMembershipLimitedToOwners": false,
        "discoverySettings": {
          "showInTeamsSearchAndSuggestions": true
        },
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
      }
    ]));
  });

  it('lists joined Microsoft Teams in the tenant', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
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
        };
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "createdDateTime": "2022-12-08T09:17:55.039Z",
          "displayName": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1" : "Team 2",
          "description": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1 description" : "Team 2 description",
          "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
          "classification": null,
          "specialization": "none",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": true
          },
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
        };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams`) {
        return {
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
        };
      }

      throw 'Invalid request';
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
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,resourceProvisioningOptions`) {
        return {
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
        };
      }
      else if ((opts.url as string).startsWith(`https://graph.microsoft.com/v1.0/teams/`)) {
        const id: string = (<string>opts.url).substring((<string>opts.url).lastIndexOf(`/`) + 1);
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
          "id": id,
          "createdDateTime": "2022-12-08T09:17:55.039Z",
          "displayName": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1" : "Team 2",
          "description": id === "02bd9fd6-8f93-4758-87c3-1fb73740a315" ? "Team 1 description" : "Team 2 description",
          "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
          "classification": null,
          "specialization": "none",
          "visibility": "public",
          "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
          "isArchived": false,
          "isMembershipLimitedToOwners": false,
          "discoverySettings": {
            "showInTeamsSearchAndSuggestions": true
          },
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
        };
      }
      else if (opts.url === `https://graph.microsoft.com/v1.0/me/joinedTeams`) {
        return {
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { output: 'json' } });
    assert(loggerLogSpy.calledWith([
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
        "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        "createdDateTime": "2022-12-08T09:17:55.039Z",
        "displayName": "Team 1",
        "description": "Team 1 description",
        "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
        "classification": null,
        "specialization": "none",
        "visibility": "public",
        "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
        "isArchived": false,
        "isMembershipLimitedToOwners": false,
        "discoverySettings": {
          "showInTeamsSearchAndSuggestions": true
        },
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
      },
      {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
        "id": "13be6971-79db-4f33-9d41-b25589ca25af",
        "createdDateTime": "2022-12-08T09:17:55.039Z",
        "displayName": "Team 2",
        "description": "Team 2 description",
        "internalId": "19:pLknmKPPkvgeaG0FtegLfjoDINeY3gvmitMkNG9H3X41@thread.tacv2",
        "classification": null,
        "specialization": "none",
        "visibility": "public",
        "webUrl": "https://teams.microsoft.com/l/team/19:a5c6eccad3fb401997756a1501d561aa%40thread.skype/conversations?groupId=8090c93e-ba7c-433e-9f39-08c7ba07c0b3&tenantId=dcd219dd-bc68-4b9b-bf0b-4a33a796be35",
        "isArchived": false,
        "isMembershipLimitedToOwners": false,
        "discoverySettings": {
          "showInTeamsSearchAndSuggestions": true
        },
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
      }
    ]));
  });
});
