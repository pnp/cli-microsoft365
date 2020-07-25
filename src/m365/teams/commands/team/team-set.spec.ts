import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.TEAMS_TEAM_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
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
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TEAM_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('sets the visibility settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          visibility: 'Public'
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', visibility: 'Public' }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the mailNickName correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          mailNickName: 'NewNickName'
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', mailNickName: 'NewNickName' }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the description settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          description: 'desc'
        })) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: true, teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', description: 'desc' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  it('sets the classification settings correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee` &&
        JSON.stringify(opts.body) === JSON.stringify({
          classification: 'MBI'
        })) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: true, teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', classification: 'MBI' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee",
            "innerError": {
              "request-id": "27b49647-a335-48f8-9a7c-f1ed9b976aaa",
              "date": "2019-04-05T12:16:48"
            }
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee', displayName: 'NewName' }
    }, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'No team found with Group Id 8231f9f2-701f-4c6e-93ce-ecb563e3c1ee');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { teamId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if visibility is not a valid visibility Private|Public', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee',
        visibility: 'hidden'
      }
    });
    assert.notStrictEqual(actual, false);
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
});