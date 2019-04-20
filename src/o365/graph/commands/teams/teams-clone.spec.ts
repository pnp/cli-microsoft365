import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-clone');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';
import request from '../../../../request';

describe(commands.TEAMS_CLONE, () => {
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
      request.post
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
    assert.equal(command.name.startsWith(commands.TEAMS_CLONE), true);
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
        assert.equal(telemetry.name, commands.TEAMS_CLONE);
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

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid GUID.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the displayName is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the partsToClone is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist"
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation on invalid visibility', () => {
    const actual = (command.validate() as CommandValidate)({ options: { visibility: 'abc' } });
    assert.notEqual(actual, true);
  });

  it('passes validation on valid \'private\' visibility', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'private'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation on valid \'public\' visibility', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: 'public'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the input is correct with mandatory parameters', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation when the input is correct with mandatory and optional parameters', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        description: "Self help community for library",
        visibility: "public",
        classification: "public"
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "abc"
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if partsToClone is set to invalid value', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "abc"
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "private"
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if visibility is set to private', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members",
        visibility: "private"
      }
    });
    assert.equal(actual, true);
  });

  it('creates a clone of a Microsoft Teams team with mandatory parameters', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
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
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, () => {
      try {

        assert(cmdInstanceLogSpy.notCalled);
        done();

      } catch (e) {
        done(e);
      }
    });
  });

  it('creates a clone of a Microsoft Teams team with mandatory parameters (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
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
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: "Library Assist",
        partsToClone: "apps,tabs,settings,channels,members"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a clone of a Microsoft Teams team with optional parameters (debug)', (done) => {
    const sinonStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone`) {
        return Promise.resolve({
          "location": "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
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
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5',
        displayName: 'Library Assist',
        partsToClone: 'apps,tabs,settings,channels,members',
        description: 'abc',
        visibility: 'public',
        classification: 'label'
      }
    }, () => {
      try {
        assert.equal(sinonStub.lastCall.args[0].url, 'https://graph.microsoft.com/v1.0/teams/15d7a78e-fd77-4599-97a5-dbb6372846c5/clone');
        assert.equal(sinonStub.lastCall.args[0].body.displayName, 'Library Assist');
        assert.equal(sinonStub.lastCall.args[0].body.partsToClone, 'apps,tabs,settings,channels,members');
        assert.equal(sinonStub.lastCall.args[0].body.description, 'abc');
        assert.equal(sinonStub.lastCall.args[0].body.visibility, 'public');
        assert.equal(sinonStub.lastCall.args[0].body.classification, 'label');
        assert.notEqual(sinonStub.lastCall.args[0].body.mailNickname.length, 0);
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
    assert(find.calledWith(commands.TEAMS_CLONE));
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