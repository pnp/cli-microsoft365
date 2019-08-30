import commands from '../../commands';
import Command, { CommandValidate, CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./teams-channel-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_CHANNEL_REMOVE, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.TEAMS_CHANNEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('passes validation when valid channelId & teamId is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.equal(actual, true);
  });

  it('fails validation if the teamId & channelId are not provided', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {

      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not provided', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not provided', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not valid channelId', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype'
      }
    });
    assert.notEqual(actual, true);
  });

  it('prompts before removing the specified channel when confirm option not passed', (done) => {
    cmdInstance.action({
      options: {
        debug: false,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
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

  it('prompts before removing the specified channel when confirm option not passed (debug)', (done) => {
    cmdInstance.action({
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
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

  it('aborts removing the specified channel when confirm option not passed and prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
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

  it('aborts removing the specified channel when confirm option not passed and prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
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

  it('removes the specified channel when prompt confirmed (debug)', (done) => {
    let teamsChannelDeleteCallIssued = false;

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
        teamsChannelDeleteCallIssued = true;
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    }, () => {
      try {
        assert(teamsChannelDeleteCallIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified channel without prompting when confirmed specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        confirm: true
      }
    }, () => {
      done();
    }, (err: any) => done(err));
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "Failed to execute Skype backend request GetThreadS2SRequest.",
            "innerError": {
              "request-id": "5a563fc6-6df2-4cd9-b0b8-9810f1110714",
              "date": "2019-08-28T19:18:30"
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
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, (err?: any) => {
      try {
        assert.equal(err.message, "Failed to execute Skype backend request GetThreadS2SRequest.");
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
    assert(find.calledWith(commands.TEAMS_CHANNEL_REMOVE));
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
