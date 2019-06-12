import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-funsettings-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_FUNSETTINGS_SET, () => {
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
    assert.equal(command.name.startsWith(commands.TEAMS_FUNSETTINGS_SET), true);
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
        assert.equal(telemetry.name, commands.TEAMS_FUNSETTINGS_SET);
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

  it('sets allowGiphy settings to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowGiphy: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: 'false' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowGiphy settings to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowGiphy: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowGiphy: 'true' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets giphyContentRating to moderate', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'moderate'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'moderate' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets giphyContentRating to strict', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            giphyContentRating: 'strict'
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', giphyContentRating: 'strict' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowStickersAndMemes to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: 'true' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowStickersAndMemes to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowStickersAndMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowStickersAndMemes: 'false' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('sets allowCustomMemes to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowCustomMemes: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'true' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCustomMemes to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'false' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCustomMemes to false (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-11f09f201302` &&
        JSON.stringify(opts.body) === JSON.stringify({
          funSettings: {
            allowCustomMemes: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-11f09f201302', allowCustomMemes: 'false' }
    }, (err?: any) => {
      try {
        assert.equal(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets fun settings of a Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315`) {
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
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315",
        allowGiphy: true,
        giphyContentRating: "moderate",
        allowStickersAndMemes: false,
        allowCustomMemes: true
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if teamId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {}
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'invalid' }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation when teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66' }
    });
    assert.equal(actual, true);
  });

  it('passes validation when allowGiphy is a valid boolean', () => {
    const actualTrue = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'true' }
    });
    const actualFalse = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'false' }
    });
    const actual = actualTrue && actualFalse;
    assert.equal(actual, true);
  });
  it('fails validation when allowGiphy is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowGiphy: 'trueish' }
    });
    assert.notEqual(actual, true);
  });
  it('passes validation when giphyContentRating is moderate or strict', () => {
    const actualModerate = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'moderate' }
    });
    const actualStrict = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'strict' }
    });
    const actual = actualModerate && actualStrict;
    assert.equal(actual, true);
  });
  it('fails validation when giphyContentRating is not moderate or strict', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', giphyContentRating: 'somethingelse' }
    });
    assert.notEqual(actual, true);
  });
  it('passes validation when allowStickersAndMemes is a valid boolean', () => {
    const actualTrue = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'true' }
    });
    const actualFalse = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'false' }
    });
    const actual = actualTrue && actualFalse;
    assert.equal(actual, true);
  });
  it('fails validation when allowStickersAndMemes is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowStickersAndMemes: 'somethingelse' }
    });
    assert.notEqual(actual, true);
  });
  it('passes validation when allowCustomMemes is a valid boolean', () => {
    const actualTrue = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'true' }
    });
    const actualFalse = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'false' }
    });
    const actual = actualTrue && actualFalse;
    assert.equal(actual, true);
  });
  it('fails validation when allowCustomMemes is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66', allowCustomMemes: 'somethingelse' }
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_FUNSETTINGS_SET));
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