import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./groupsetting-remove');

describe(commands.GROUPSETTING_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
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
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });
    promptOptions = undefined;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.delete,
      global.setTimeout,
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      fs.readFileSync,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GROUPSETTING_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('removes the specified group setting without prompting for confirmation when confirm option specified', (done) => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groupSettings/28beab62-7540-4db1-a23f-29a6018a3848') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848', confirm: true } }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the specified group setting without prompting for confirmation when confirm option specified (debug)', (done) => {
    const deleteRequestStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groupSettings/28beab62-7540-4db1-a23f-29a6018a3848') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848', confirm: true } }, () => {
      try {
        assert(deleteRequestStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before removing the specified group setting when confirm option not passed', (done) => {
    command.action(logger, { options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
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

  it('prompts before removing the specified group setting when confirm option not passed (debug)', (done) => {
    command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
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

  it('aborts removing the group setting when prompt not confirmed', (done) => {
    const postSpy = sinon.spy(request, 'delete');

    command.action(logger, { options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts removing the group setting when prompt not confirmed (debug)', (done) => {
    const postSpy = sinon.spy(request, 'delete');

    command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(postSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the group setting when prompt confirmed', (done) => {
    const postStub = sinon.stub(request, 'delete').callsFake(() => Promise.resolve());

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(postStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the group setting when prompt confirmed (debug)', (done) => {
    const postStub = sinon.stub(request, 'delete').callsFake(() => Promise.resolve());

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    });
    command.action(logger, { options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(postStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when group setting is not found', (done) => {
    sinon.stub(request, 'delete').callsFake(() => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'File Not Found.' } } } });
    });

    command.action(logger, { options: { debug: false, confirm: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('File Not Found.')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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

  it('supports specifying id', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('supports specifying confirmation flag', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--confirm') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', async () => {
    const actual = await command.validate({ options: { id: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', async () => {
    const actual = await command.validate({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});