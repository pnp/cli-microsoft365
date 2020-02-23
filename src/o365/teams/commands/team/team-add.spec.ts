import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_TEAM_ADD, () => {
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
      request.post,
      request.put
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.TEAMS_TEAM_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {  
        groupId:'f9526e6a-1d0d-4421-8882-88a70975a00c'
      }
    }, () => {
      done();
    }, (err: any) => done(err));
  });

  it('correctly handles error when creating a Team', (done) => {
    sinon.stub(request, 'put').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {  
        groupId:'f9526e6a-1d0d-4421-8882-88a70975a00c'
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
    assert(find.calledWith(commands.TEAMS_TEAM_ADD));
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