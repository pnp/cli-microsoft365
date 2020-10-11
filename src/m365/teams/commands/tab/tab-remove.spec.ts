import * as assert from 'assert';
import * as chalk from 'chalk';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./tab-remove');

describe(commands.TEAMS_TAB_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
    promptOptions = undefined;
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TAB_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid channelId, teamId and tabId is specified', () => {
    const actual = command.validate({
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId , channelId and tabId are not provided', (done) => {
    const actual = command.validate({
      options: {

      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not valid channelId', (done) => {
    const actual = command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = command.validate({
      options: {
        teamId: 'invalid',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.notStrictEqual(actual, true);
  });
  it('fails validation if the tabId is not a valid guid', () => {
    const actual = command.validate({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        tabId: 'invalid',
      }
    });
    assert.notStrictEqual(actual, true);
  });


  it('prompts before removing the specified tab when confirm option not passed', (done) => {
    command.action(logger, {
      options: {
        debug: false,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
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

  it('prompts before removing the specified tab when confirm option not passed (debug)', (done) => {
    command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
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

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the specified tab when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified tab by id when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    Utils.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
      try {
        assert(loggerSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });


  it('removes the specified tab without prompting when confirmed specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if ((opts.url as string).indexOf(`tabs/d66b8110-fcad-49e8-8159-0d488ddb7656`) > -1) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: '00000000-0000-0000-0000-000000000000',
        tabId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        confirm: true
      }
    }, () => {
      assert(loggerSpy.calledWith(chalk.green('DONE')));
      done();
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
