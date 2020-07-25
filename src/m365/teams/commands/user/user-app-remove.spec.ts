import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./user-app-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.TEAMS_USER_APP_REMOVE, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_USER_APP_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the userId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        userId: 'invalid',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        userId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('prompts before removing the app when confirmation argument is not passed', (done) => {
    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
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

  it('aborts removing the app when confirmation prompt is not continued', (done) => {
    const requestDeleteSpy = sinon.stub(request, 'delete');

    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY='
      }
    }, () => {
      try {
        assert(requestDeleteSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the app for the specified user when confirmation is specified (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.toString().indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        debug: true,
        confirm: true
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

  it('removes the app for the specified user when prompt is confirmed (debug)', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.toString().indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };
    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        debug: true
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

  it('correctly handles error while removing teams app', (done) => {
    const error = 'An error has occurred';
    sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url.toString().indexOf(`/users/c527a470-a882-481c-981c-ee6efaba85c7/teamwork/installedApps/YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=`) > -1) {
        return Promise.reject(error);
      }
      return Promise.reject('Invalid Request');
    });

    cmdInstance.action({
      options: {
        userId: 'c527a470-a882-481c-981c-ee6efaba85c7',
        appId: 'YzUyN2E0NzAtYTg4Mi00ODFjLTk4MWMtZWU2ZWZhYmE4NWM3IyM0ZDFlYTA0Ny1mMTk2LTQ1MGQtYjJlOS0wZDI4NTViYTA1YTY=',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)));
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