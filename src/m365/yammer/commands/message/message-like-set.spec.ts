import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';

const command: Command = require('./message-like-set');
describe(commands.MESSAGE_LIKE_SET, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      request.post,
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
    assert.strictEqual(command.name.startsWith(commands.MESSAGE_LIKE_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
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

  it('passes validation if enabled set to "true"', async () => {
    const actual = await command.validate({ options: { id: 10123123, enable: 'true' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if enabled set to "true"', async () => {
    const actual = await command.validate({ options: { id: 10123123, enable: 'false' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('does not pass validation if enable not set to either "true" or "false"', async () => {
    const actual = await command.validate({ options: { id: 10123123, enable: 'fals' } }, commandInfo);
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

  it('prompts when confirmation argument not passed', (done) => {
    command.action(logger, { options: { debug: false, id: 1231231, enable: 'false' } }, () => {
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

  it('calls the service when liking a message', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 1231231 } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when liking a message and confirm passed', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 1231231, confirm: 'true' } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when liking a message and enabled set to true', (done) => {
    const requestPostedStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 1231231, enable: 'true' } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service when disliking a message and confirming', (done) => {
    const requestPostedStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: 1231231, enable: 'false', confirm: true } }, () => {
      try {
        assert(requestPostedStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts when disliking and confirmation parameter is denied', (done) => {
    command.action(logger, { options: { debug: false, id: 1231231, enable: 'false', confirm: false } }, () => {
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

  it('calls the service when disliking a message and confirmation is hit', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/liked_by/current.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });

    command.action(logger, { options: { debug: true, id: 1231231, enable: 'false' } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Aborts execution when enabled set to false and confirmation is not given', (done) => {
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });
    command.action(logger, { options: { debug: false, id: 1231231, enable: 'false' } }, () => {
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