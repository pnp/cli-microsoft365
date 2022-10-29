import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
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
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
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

    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: true };
    });
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      Cli.prompt,
      (command as any).exec,
      (process as any).kill
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent,
      pid.getProcessName,
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

  it('prompts before clear cache when confirm option not passed', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async (options) => {
      promptOptions = options;
      return { continue: false };
    });

    await command.action(logger, {
      options: {}
    });
    let promptIssued = false;

    if (promptOptions && promptOptions.type === 'confirm') {
      promptIssued = true;
    }
    assert(promptIssued);
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

  it('fails to remove teams cache when exec fails randomly when killing teams.exe process', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const error = new Error('random error');
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'taskkill /IM "Teams.exe" /F') {
        throw error;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { confirm: true } } as any), new CommandError('random error'));
  });

  it('fails to remove teams cache when exec fails randomly when removing cache folder', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const error = new Error('random error');
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'taskkill /IM "Teams.exe" /F') {
        return { stdout: '' };
      }
      if (opts === 'cd %userprofile% && rmdir /s /q AppData\\Roaming\\Microsoft\\Teams') {
        throw error;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { confirm: true } } as any), new CommandError('random error'));
  });

  it('removes Teams cache from macOs platform without prompting.', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: '' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes teams cache when teams is currently not active', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const error = new Error('ERROR: The process "Teams.exe" not found.');
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'taskkill /IM "Teams.exe" /F') {
        throw error;
      }
      if (opts === 'cd %userprofile% && rmdir /s /q AppData\\Roaming\\Microsoft\\Teams') {
        return;
      }
      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes Teams cache from win32 platform without prompting.', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: '' });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: {
        confirm: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes Teams cache from darwin platform with prompting.', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: 'pid' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    await command.action(logger, {
      options: {
        debug: true
      }
    });
    assert(true);
  });

  it('aborts cache clearing when no cache folder is found', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(fs, 'existsSync').callsFake(() => false);

    await command.action(logger, {
      options: {
        verbose: true
      }
    });
  });

  it('aborts cache clearing from Teams when prompt not confirmed', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    const postSpy = sinon.spy(request, 'delete');
    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'prompt').callsFake(async () => (
      { continue: false }
    ));

    await command.action(logger, { options: {} });
    assert(postSpy.notCalled);
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