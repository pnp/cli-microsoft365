import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-remove');

describe(commands.TEAMS_APP_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    requests = [];
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.delete,
      Cli.prompt
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
    const actual = command.validate({
      options: { id: 'invalid' }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const actual = command.validate({
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

    command.action(logger, { options: { debug: false, id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } }, () => {
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

    command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } }, () => {
      try {
        assert(removeTeamsAppCalled);
        done();
      }
      catch (e) {
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

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } }, () => {
      try {
        assert(removeTeamsAppCalled);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing Teams app when prompt not confirmed', (done) => {
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22` } }, () => {
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

    command.action(logger, { options: { debug: false, filePath: 'teamsapp.zip', id: `e3e29acb-8c79-412b-b746-e6c39ff4cd22`, confirm: true } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});