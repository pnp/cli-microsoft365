import commands from '../../commands';
import Command, { CommandError, CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-message-get');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_MESSAGE_GET, () => {
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
      commandWrapper: {
        command: commands.TEAMS_MESSAGE_GET
      },
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    auth.service = new Service('https://graph.microsoft.com');
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      global.setTimeout
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
    assert.equal(command.name.startsWith(commands.TEAMS_MESSAGE_GET), true);
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
        assert.equal(telemetry.name, commands.TEAMS_MESSAGE_GET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts when not logged in to the Microsoft Graph', (done) => {
    auth.service = new Service('https://graph.microsoft.com');
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

  it('defines alias', () => {
    const alias = command.alias();
    assert.notEqual(typeof alias, 'undefined');
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
    assert(find.calledWith(commands.TEAMS_MESSAGE_GET));
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

  it('fails validation if teamId, channelId and messageId are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if channelId and messageId are not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if messageId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "5f5d7b71-1161-44",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        messageId: "1540911392778"
      }
    });
    assert.notEqual(actual, true);
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

  it('validates for a correct input', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        messageId: "1540911392778"
      }
    });
    assert.equal(actual, true);
  });

  it('retrieves the specified message (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/5f5d7b71-1161-44d8-bcc1-3da710eb4171/channels/19:88f7e66a8dfe42be92db19505ae912a8@thread.skype/messages/1540911392778`) {
        return Promise.resolve({
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        });
      }

      return Promise.reject('Invalid Request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        messageId: "1540911392778"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        }));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('retrieves the specified message', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams/5f5d7b71-1161-44d8-bcc1-3da710eb4171/channels/19:88f7e66a8dfe42be92db19505ae912a8@thread.skype/messages/1540911392778`) {
        return Promise.resolve({
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        });
      }

      return Promise.reject('Invalid Request');
    });

    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: "5f5d7b71-1161-44d8-bcc1-3da710eb4171",
        channelId: "19:88f7e66a8dfe42be92db19505ae912a8@thread.skype",
        messageId: "1540911392778"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          attachments: [],
          body: { "contentType": "text", "content": "Konnichiwa" },
          createdDateTime: "2018-10-28T15:56:25.116Z",
          deleted: false,
          etag: "1540742185116",
          from: { "application": null, "device": null, "user": { "id": "c500ecce-645d-4fe1-a2ea-b70f32416b51", "displayName": "Arjen Bloemsma", "identityProvider": "Aad" } },
          id: "1540742185116",
          importance: "normal",
          lastModifiedDateTime: null,
          locale: "en-us",
          mentions: [],
          messageType: "message",
          policyViolation: null,
          reactions: [],
          replyToId: null,
          subject: "",
          summary: null
        }));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service('https://graph.microsoft.com');
    auth.service.connected = true;
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