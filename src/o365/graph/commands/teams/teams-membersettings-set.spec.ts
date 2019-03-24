import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-membersettings-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_MEMBERSETTINGS_SET, () => {
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
    assert.equal(command.name.startsWith(commands.TEAMS_MEMBERSETTINGS_SET), true);
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
        assert.equal(telemetry.name, commands.TEAMS_MEMBERSETTINGS_SET);
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

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('sets the allowAddRemoveApps setting to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowAddRemoveApps: true
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
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'true' }
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

  it('sets the allowAddRemoveApps setting to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowAddRemoveApps: false
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
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'false' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCreateUpdateChannels, allowCreateUpdateRemoveConnectors and allowDeleteChannels to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowCreateUpdateChannels: true,
            allowCreateUpdateRemoveConnectors: true,
            allowDeleteChannels: true
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
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'true', allowCreateUpdateRemoveConnectors: 'true', allowDeleteChannels: 'true' }
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

  it('sets allowCreateUpdateChannels, allowCreateUpdateRemoveTabs and allowDeleteChannels to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowCreateUpdateChannels: false,
            allowCreateUpdateRemoveTabs: false,
            allowDeleteChannels: false
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
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'false', allowCreateUpdateRemoveTabs: 'false', allowDeleteChannels: 'false' }
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

  it('fails validation if the teamId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { teamId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } });
    assert.equal(actual, true);
  });

  it('fails validation if allowAddRemoveApps is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateChannels is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveConnectors is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveTabs is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation if allowDeleteChannels is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'invalid'
      }
    });
    assert.notEqual(actual, true);
  });

  it('passes validation if allowAddRemoveApps is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'false'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowAddRemoveApps is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'true'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'false'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'true'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'false'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'true'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'false'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'true'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowDeleteChannels is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'false'
      }
    });
    assert.equal(actual, true);
  });

  it('passes validation if allowDeleteChannels is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'true'
      }
    });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.TEAMS_MEMBERSETTINGS_SET));
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