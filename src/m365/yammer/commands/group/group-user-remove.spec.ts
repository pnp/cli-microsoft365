import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./group-user-remove');

describe(commands.GROUP_USER_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let requests: any[];
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUP_USER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation with parameters', async () => {
    const actual = await command.validate({ options: { id: 10123123 } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('id must be a number', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('userId must be a number', async () => {
    const actual = await command.validate({ options: { id: 10, userId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('calls the service if the current user is removed from the group', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, id: 1231231 } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service if the user 989998789 is removed from the group 1231231 with the confirm command', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 1231231, userId: 989998789, confirm: true } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service if the user 989998789 is removed from the group 1231231', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, id: 1231231, userId: 989998789 } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removal when confirmation argument not passed', (done) => {
    const promptStub: sinon.SinonStub = sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, id: 1231231, userId: 989998789 } }, () => {
      try {
        assert(promptStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts execution when prompt not confirmed', (done) => {
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, id: 1231231, userId: 989998789 } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 