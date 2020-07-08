import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./channel-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_CHANNEL_LIST, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TEAMS_CHANNEL_LIST), true);
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

  it('correctly handles error when retrieving all teams', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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
});