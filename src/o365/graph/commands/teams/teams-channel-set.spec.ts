import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-channel-set');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_CHANNEL_SET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
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
    auth.service = new Service('https://graph.microsoft.com');
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.patch
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
    assert.equal(command.name.startsWith(commands.TEAMS_CHANNEL_SET), true);
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
        assert.equal(telemetry.name, commands.TEAMS_CHANNEL_SET);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('correctly validates the arguments', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelName: 'Reviews',
        newChannelName: 'Gen',
        description: 'this is a new description'
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        channelName: 'Reviews',
        newChannelName: 'Gen',
        description: 'this is a new description'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelName: 'Reviews',
        newChannelName: 'Gen',
        description: 'this is a new description'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation when no channelName is specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        newChannelName: 'Reviews',
        description: 'this is a new description'
      }
    });
    assert.notEqual(actual, true);
    done();
  });
  it('fails validation when no newChannelName is specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelName: 'Reviews',
        description: 'this is a new description'
      }
    });
    assert.notEqual(actual, true);
    done();
  });
  it('fails validation when channelName is General', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelName: 'General',
        newChannelName: 'Reviews',
        description: 'this is a new description'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('correctly patches channel updates for the Microsoft Teams team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          value: [
            { "id": "19:3b9de2fb06f145609c42ede0d5d305fb@thread.skype", "displayName": "General", "description": "New group for Policy Doco" },
            { "id": "19:a3a51b130bb34dba9712ef9675bbb504@thread.skype", "displayName": "Discussion", "description": "This is a channel for Discussion for the document Policy Doco.docx " },
            { "id": "19:01740d80914145dca1a47f3f7ea1f78d@thread.skype", "displayName": "bla", "description": "This is a channel for Review for the document Policy Doco.docx " }]
        });
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:01740d80914145dca1a47f3f7ea1f78d@thread.skype` &&
        JSON.stringify(opts.body) === `{"description":"new social","displayName":"bla2"}`
      ) {
        return Promise.resolve({});
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
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: "bla",
        newChannelName: "bla2",
        description: "new social"
      }
    }, (err?: any) => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly patches channel updates for the Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          value: [
            { "id": "19:3b9de2fb06f145609c42ede0d5d305fb@thread.skype", "displayName": "General", "description": "New group for Policy Doco" },
            { "id": "19:a3a51b130bb34dba9712ef9675bbb504@thread.skype", "displayName": "Discussion", "description": "This is a channel for Discussion for the document Policy Doco.docx " },
            { "id": "19:01740d80914145dca1a47f3f7ea1f78d@thread.skype", "displayName": "bla", "description": "This is a channel for Review for the document Policy Doco.docx " }]
        });
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels/19:01740d80914145dca1a47f3f7ea1f78d@thread.skype` &&
        JSON.stringify(opts.body) === `{"description":"new social","displayName":"bla2"}`
      ) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: "bla",
        newChannelName: "bla2",
        description: "new social"
      }
    }, (err?: any) => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('fails to patch channel updates for the Microsoft Teams team when channel does not exists', (done) => {

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          "value": [
            { "id": "19:3b9de2fb06f145609c42ede0d5d305fb@thread.skype", "displayName": "General", "description": "New group for Policy Doco" },
            { "id": "19:a3a51b130bb34dba9712ef9675bbb504@thread.skype", "displayName": "Discussion", "description": "This is a channel for Discussion for the document Policy Doco.docx " },
            { "id": "19:01740d80914145dca1a47f3f7ea1f78d@thread.skype", "displayName": "Test", "description": "This is a channel for Review for the document Policy Doco.docx " }]
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
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: "bla",
        newChannelName: "bla2",
        description: "new social"
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('fails to patch channel updates for the Microsoft Teams team when channel information is incorrect', (done) => {

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/channels`) {
        return Promise.resolve({
          "value": [
            { "id1": "19:3b9de2fb06f145609c42ede0d5d305fb@thread.skype", "displayName1": "General", "description": "New group for Policy Doco" },
            { "id1": "19:a3a51b130bb34dba9712ef9675bbb504@thread.skype", "displayName1": "Discussion", "description": "This is a channel for Discussion for the document Policy Doco.docx " },
            { "id1": "19:01740d80914145dca1a47f3f7ea1f78d@thread.skype", "displayName1": "Test", "description": "This is a channel for Review for the document Policy Doco.docx " }]
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
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: "bla",
        newChannelName: "bla2",
        description: "new social"
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
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
    assert(find.calledWith(commands.TEAMS_CHANNEL_SET));
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