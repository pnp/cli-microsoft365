import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./user-app-add');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_USER_APP_ADD, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_USER_APP_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userId: 'invalid',
        appId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the appId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        appId: 'not-c49b-4fd4-8223-28f0ac3a6402',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        appId: '15d7a78e-fd77-4599-97a5-dbb6372846c6',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('adds app from the catalog for the specified user', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.body) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/beta/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('adds app from the catalog for the specified user (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps` &&
        JSON.stringify(opts.body) === `{"teamsApp@odata.bind":"https://graph.microsoft.com/beta/appCatalogs/teamsApps/4440558e-8c73-4597-abc7-3644a64c4bce"}`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: '4440558e-8c73-4597-abc7-3644a64c4bce',
        debug: true
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

  it('correctly handles error while installing teams app', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: '4440558e-8c73-4597-abc7-3644a64c4bce'
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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