import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./cache-remove');

describe(commands.CACHE_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    sinon.stub(Cli.getInstance().config, 'all').value({});
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

    promptOptions = undefined;

    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: true });
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.prompt
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      Cli.getInstance().config.all
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.CACHE_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before clear cache when confirm option not passed', (done) => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      promptOptions = options;
      cb({ continue: false });
    });

    command.action(logger, {
      options: {}
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

  it('fails validation if called from docker container.', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': 'docker' });

    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if not called from win32 or darwin platform.', async () => {
    sinon.stub(process, 'platform').value('android');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if called from win32 or darwin platform.', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    const actual = await command.validate({
      options: {}
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails to remove teams cache when exec fails', (done) => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    const error = new Error('ERROR: The process "Teams.exe" not found.');
    const exec = sinon.stub(command, 'exec' as any).throws(error);

    command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(`ERROR: The process "Teams.exe" not found.`));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        exec.restore();
      }
    });
  });

  it('fails to remove teams cache when exec fails randomly', (done) => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    const error = new Error('random error');
    const exec = sinon.stub(command, 'exec' as any).throws(error);

    command.action(logger, {
      options: {
        confirm: true
      }
    }, (err) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(`random error`));
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        exec.restore();
      }
    });
  });

  it('removes Teams cache from macOs platform without prompting.', (done) => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    const exec = sinon.stub(command, 'exec' as any).returns({ stdout: 'pid' });
    const kill = sinon.stub(process, 'kill' as any).returns(null);

    command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    }, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        exec.restore();
        kill.restore();
      }
    });
  });

  it('removes Teams cache from win32 platform without prompting.', (done) => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    const exec = sinon.stub(command, 'exec' as any).returns({ stdout: '' });

    command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    }, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        exec.restore();
      }
    });
  });

  it('removes Teams cache from darwin platform with prompting.', (done) => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    const exec = sinon.stub(command, 'exec' as any).returns({ stdout: 'pid' });
    const kill = sinon.stub(process, 'kill' as any).returns(null);

    command.action(logger, {
      options: {
        debug: true
      }
    }, () => {
      try {
        assert(true);
        done();
      }
      catch (e) {
        done(e);
      }
      finally {
        exec.restore();
        kill.restore();
      }
    });
  });

  it('aborts cache clearing from Teams when prompt not confirmed', (done) => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake((options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    });

    command.action(logger, { options: {} }, () => {
      try {
        assert(postSpy.notCalled);
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
});