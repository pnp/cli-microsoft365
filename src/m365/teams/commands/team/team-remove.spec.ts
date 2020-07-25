import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./team-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_TEAM_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let promptOptions: any;

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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    promptOptions = undefined;
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.delete
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TEAM_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '61703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('passes validation when valid teamId is specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('prompts before removing the specified team when confirm option not passed', (done) => {
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000"} }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified team when confirm option not passed (debug)', (done) => {
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified team when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified team when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified team when prompt confirmed (debug)', (done) => {
    let teamsDeleteCallIssued = false;
   
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
        teamsDeleteCallIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(teamsDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified team without prompting when confirmed specified', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", confirm: true } }, () => {
      done();
    }, (err: any) => done(err));
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/8231f9f2-701f-4c6e-93ce-ecb563e3c1ee`) {
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

    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: { teamId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
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