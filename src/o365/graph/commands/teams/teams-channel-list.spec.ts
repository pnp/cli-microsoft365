import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-channel-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_CHANNEL_LIST, () => {
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
    assert.equal(command.name.startsWith(commands.TEAMS_CHANNEL_LIST), true);
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
        assert.equal(telemetry.name, commands.TEAMS_CHANNEL_LIST);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '00000000-0000-0000-0000-000000000000',
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('correctly lists all channels in a Microsoft teams team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          value: [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
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
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social"

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

  it('correctly lists all channels in a Microsoft teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          value: [{ "id": "19:17de660d16844149ab3f0240405f9316@thread.skype", "displayName": "General", "description": "Test group for office cli commands", "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype", "displayName": "Development", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }, { "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype", "displayName": "Social", "description": null, "isFavoriteByDefault": null, "email": "", "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8" }]
        });
      }
      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith([
          {
            "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
            "displayName": "General"
          },
          {
            "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
            "displayName": "Development"
          },
          {
            "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
            "displayName": "Social"
          }
        ]));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('outputs all data in json output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          value: [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"
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
    cmdInstance.action({
      options: {
        debug: false,
        output: 'json',
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          [
            {
              "id": "19:17de660d16844149ab3f0240405f9316@thread.skype",
              "displayName": "General",
              "description": "Test group for office cli commands",
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a17de660d16844149ab3f0240405f9316%40thread.skype/General?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"

            },
            {
              "id": "19:e14b10cd0b684901b53d14e89aa4221f@thread.skype",
              "displayName": "Development",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3ae14b10cd0b684901b53d14e89aa4221f%40thread.skype/Development?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"

            },
            {
              "id": "19:12ff25ec5325468dba1f73522cd08248@thread.skype",
              "displayName": "Social",
              "description": null,
              "isFavoriteByDefault": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19%3a12ff25ec5325468dba1f73522cd08248%40thread.skype/Social?teamId=290a87a4-38f4-4f6c-a664-9dddf09bdbcc&tenantId=3a7a651b-2620-433b-a1a3-42de27ae94e8"

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
    assert(find.calledWith(commands.TEAMS_CHANNEL_LIST));
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