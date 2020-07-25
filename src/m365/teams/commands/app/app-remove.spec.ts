import * as sinon from 'sinon';
import * as assert from 'assert';
import appInsights from '../../../../appInsights';
import request from '../../../../request';
import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import auth from '../../../../Auth';
const command: Command = require('./app-remove');
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.TEAMS_APP_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let requests: any[];

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
        cb({ continue: false });
      }
    };
    requests = [];
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the id is not a valid GUID.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: { id: 'invalid' }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22"
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('remove Teams app in the tenant app catalog with confirmation', (done) => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } }, () => {
      try {
        assert(removeTeamsAppCalled);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('remove Teams app in the tenant app catalog with confirmation (debug)', (done) => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } }, () => {
      try {
        assert(removeTeamsAppCalled);
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('remove Teams app in the tenant app catalog without confirmation', (done) => {
    let removeTeamsAppCalled = false;
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/e3e29acb-8c79-412b-b746-e6c39ff4cd22`) {
        removeTeamsAppCalled = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');      
    });

    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({ options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } }, () => {
      try {
        assert(removeTeamsAppCalled);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing Teams app when prompt not confirmed', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when removing app', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      } catch (e) {
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