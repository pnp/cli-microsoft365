import commands from '../../commands';
import Command, { CommandValidate, CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./channel-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.TEAMS_CHANNEL_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let promptOptions: any;

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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_CHANNEL_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes validation when valid channelId & teamId is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation when channelName & teamId is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelName: 'Channel Name',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId & channelId are not provided', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {

      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId & channelName are not provided', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {

      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the channelId is not valid channelId', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if both channelName and channelId are provided', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelId: '19:f3dcbb1674574677abcae89cb626f1e6@thread.skype',
        channelName: 'channelname'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails to remove channel when channel does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'channelName'`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656',
        channelName: 'channelName',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('removes the specified channel by id when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {

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
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
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

  it('removes the specified channel by name when prompt confirmed (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'channelName'`) > -1) {
        return Promise.resolve({
          value: [
            {
              "id": "19:f3dcbb1674574677abcae89cb626f1e6@thread.skype",
              "displayName": "channelName",
              "description": null,
              "email": "",
              "webUrl": "https://teams.microsoft.com/l/channel/19:f3dcbb1674574677abcae89cb626f1e6%40thread.skype/%F0%9F%92%A1+Ideas?groupId=d66b8110-fcad-49e8-8159-0d488ddb7656&tenantId=eff8592e-e14a-4ae8-8771-d96d5c549e1c"
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/d66b8110-fcad-49e8-8159-0d488ddb7656/channels/19%3Af3dcbb1674574677abcae89cb626f1e6%40thread.skype`) {
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
        channelName: 'channelName',
        teamId: 'd66b8110-fcad-49e8-8159-0d488ddb7656'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
      assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
        assert.strictEqual(err.message, "Failed to execute Skype backend request GetThreadS2SRequest.");
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
