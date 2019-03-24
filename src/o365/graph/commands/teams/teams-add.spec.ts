import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./teams-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.TEAMS_ADD, () => {
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
      request.post,
      request.put
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
    assert.equal(command.name.startsWith(commands.TEAMS_ADD), true);
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
        assert.equal(telemetry.name, commands.TEAMS_ADD);
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

  it('fails validation if the groupId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '61703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the groupId and name are specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '61703ac8a-c49b-4fd4-8223-80c3',
        name:'Architecture'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the groupId and description are specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '61703ac8a-c49b-4fd4-8223-80c3',
        description:'Architecture'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('passes validation if the groupId is a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '25bd7c99-619a-e411-80c3-a0d3c1f2861f'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('passes validation if the name and description exist with blank groupId', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture',
        description:'Architecture Discussion'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('fails validation if the groupId and name are blank.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        description: 'Architecture'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the groupId and description are blank.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        name: 'Architecture'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('creates Microsoft Teams team in the tenant (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({
          headers : {
            location: "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
          }
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
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('f9526e6a-1d0d-4421-8882-88a70975a00c'));
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team in the tenant', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({
          headers : {
            location: "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
          }
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
        name: 'Architecture',
        description: 'Architecture Discussion'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('f9526e6a-1d0d-4421-8882-88a70975a00c'));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('creates Microsoft Teams team for a group in the tenant (debug)', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/0ee1db97-37e3-4223-a44f-f389400ad1a0/team`) {
        return Promise.resolve({
          headers : {
            location: "https://api.teams.skype.com/beta/groups('0ee1db97-37e3-4223-a44f-f389400ad1a0')/team('0ee1db97-37e3-4223-a44f-f389400ad1a0')"
          }
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
        groupId:'0ee1db97-37e3-4223-a44f-f389400ad1a0'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('0ee1db97-37e3-4223-a44f-f389400ad1a0'));
        assert(cmdInstanceLogSpy.calledWith(vorpal.chalk.green('DONE')));

        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates Microsoft Teams team for a group in the tenant', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/f9526e6a-1d0d-4421-8882-88a70975a00c/team`) {
        return Promise.resolve({
          headers : {
            location: "/teams('f9526e6a-1d0d-4421-8882-88a70975a00c')/operations('6cf64f96-08c3-4173-9919-eaf7684aae9a')"
          }
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
        groupId:'f9526e6a-1d0d-4421-8882-88a70975a00c'
      }
    }, () => {
      try {
      
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
    assert(find.calledWith(commands.TEAMS_ADD));
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