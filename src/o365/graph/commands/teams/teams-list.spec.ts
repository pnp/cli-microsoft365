import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-list');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service();
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TEAMS_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.TEAMS_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not connected to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Connect to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft Teams in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2"
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "description": "Team 3 description",
              "displayName": "Team 3"
            }
          ]
        });
      } else if (opts.url.startsWith(`https://graph.microsoft.com/beta/teams/`)) {
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
        })
      };

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
          },
          {
            "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
            "displayName": "Team 3",
            "description": "Team 3 description",
            "isArchived": false
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly handles when listing a team a user is not member of', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            }
          ]
        });
      } else if (opts.url === `https://graph.microsoft.com/beta/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315`) {
        return Promise.reject({ statusCode: 403 });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
            "displayName": "Team 1",
            "description": "Team 1 description",
            "isArchived": undefined
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly handles when listing a team a user is not member of, and the graph returns an unexpected error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Invalid request')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists Microsoft Teams in the tenant (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2"
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "description": "Team 3 description",
              "displayName": "Team 3"
            }
          ]
        });
      } else if (opts.url.startsWith(`https://graph.microsoft.com/beta/teams/`)) {
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
        })
      };

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
          },
          {
            "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
            "displayName": "Team 3",
            "description": "Team 3 description",
            "isArchived": false
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists joined Microsoft Teams in the tenant', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2"
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "description": "Team 3 description",
              "displayName": "Team 3"
            }
          ]
        });
      } else if (opts.url.startsWith(`https://graph.microsoft.com/beta/teams/`)) {
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
        })
      } else if (opts.url === `https://graph.microsoft.com/beta/me/joinedTeams`) {
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
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "displayName": "Team 3",
              "description": "Team 3 description",
              "isArchived": false
            }
          ]
        });
      };

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { joined: true, debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
          },
          {
            "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
            "displayName": "Team 3",
            "description": "Team 3 description",
            "isArchived": false
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName,description`) {
        return Promise.resolve({
          "value": [
            {
              "id": "02bd9fd6-8f93-4758-87c3-1fb73740a315",
              "description": "Team 1 description",
              "displayName": "Team 1"
            },
            {
              "id": "13be6971-79db-4f33-9d41-b25589ca25af",
              "description": "Team 2 description",
              "displayName": "Team 2"
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "description": "Team 3 description",
              "displayName": "Team 3"
            }
          ]
        });
      } else if (opts.url.startsWith(`https://graph.microsoft.com/beta/teams/`)) {
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
        })
      } else if (opts.url === `https://graph.microsoft.com/beta/me/joinedTeams`) {
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
            },
            {
              "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
              "displayName": "Team 3",
              "description": "Team 3 description",
              "isArchived": false
            }
          ]
        });
      };

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
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
          },
          {
            "id": "8090c93e-ba7c-433e-9f39-08c7ba07c0b3",
            "displayName": "Team 3",
            "description": "Team 3 description",
            "isArchived": false
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_LIST));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});