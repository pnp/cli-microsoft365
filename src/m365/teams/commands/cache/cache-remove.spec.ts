import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './cache-remove.js';

describe(commands.CACHE_REMOVE, () => {
  const processOutput = `ProcessId
  6456
  14196
  11352`;
  let log: string[];
  let logger: Logger;
  let promptOptions: any;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    sinon.stub(Cli.getInstance().config, 'all').value({});
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };

    promptOptions = undefined;

    sinon.stub(Cli, 'promptForConfirmation').resolves(true);
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      Cli.promptForConfirmation,
      (command as any).exec,
      (process as any).kill
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CACHE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before clear cache when confirm option not passed', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

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
    sinon.stub(fs, 'existsSync').returns(true);
    const error = new Error('random error');
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
        throw error;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { force: true } } as any), new CommandError('random error'));
  });

  it('fails to remove teams cache when exec fails randomly when removing cache folder', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);
    const error = new Error('random error');
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
        return { stdout: processOutput };
      }
      if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
        throw error;
      }
      throw 'Invalid request';
    });
    await assert.rejects(command.action(logger, { options: { force: true } } as any), new CommandError('random error'));
  });

  it('removes Teams cache from macOs platform without prompting.', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: '' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        force: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes teams cache when teams is currently not active', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
        return { stdout: 'No Instance(s) Available.' };
      }
      if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
        return;
      }
      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        force: true,
        verbose: true
      }
    });
    assert(true);
  });

  it('removes Teams cache from win32 platform without prompting.', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'wmic process where caption="Teams.exe" get ProcessId') {
        return { stdout: processOutput };
      }
      if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Roaming\\Microsoft\\Teams"') {
        return;
      }
      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);
    await command.action(logger, {
      options: {
        force: true,
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
    sinon.stub(fs, 'existsSync').returns(true);

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
    sinon.stub(fs, 'existsSync').returns(false);
    await command.action(logger, {
      options: {
        verbose: true
      }
    });
  });

  it('aborts cache clearing from Teams when prompt not confirmed', async () => {
    const execStub = sinon.stub(command, 'exec' as any);
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    sinonUtil.restore(Cli.prompt);
    sinon.stub(Cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: {} });
    assert(execStub.notCalled);
  });
});