import assert from 'assert';
import fs from 'fs';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './cache-remove.js';
import os, { homedir } from 'os';

describe(commands.CACHE_REMOVE, () => {
  const processOutput = `ProcessId
  6456
  14196
  11352`;
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    sinon.stub(cli.getConfig(), 'all').value({});
    commandInfo = cli.getCommandInfo(command);
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
    loggerLogSpy = sinon.spy(logger, 'log');

    sinon.stub(cli, 'promptForConfirmation').resolves(true);
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      cli.promptForConfirmation,
      (command as any).exec,
      (process as any).kill
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.CACHE_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before clear cache when force option not passed', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });

    sinonUtil.restore(cli.promptForConfirmation);
    const confirmationStub = sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, {
      options: {}
    });
    assert(confirmationStub.calledOnce);
  });

  it('fails validation if client is not a valid client option', async () => {
    sinon.stub(process, 'platform').value('win32');
    const actual = await command.validate({
      options: {
        client: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
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

  it('fails to remove teams cache when exec fails randomly when killing teams.exe process using classic client', async () => {
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
    await assert.rejects(command.action(logger, { options: { client: 'classic', force: true } } as any), new CommandError('random error'));
  });

  it('fails to remove teams cache when exec fails randomly when removing cache folder using classic client', async () => {
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
    await assert.rejects(command.action(logger, { options: { client: 'classic', force: true } } as any), new CommandError('random error'));
  });

  it('shows error message when exec fails when removing the teams cache folder on mac os', async () => {
    const deleteError = {
      code: 1,
      killed: false,
      signal: null,
      cmd: 'rm -r "/Users/John/Library/Group Containers/UBF8T346G9.com.microsoft.teams"',
      stdout: '',
      stderr: 'rm: /Users/John/Library/Group Containers/UBF8T346G9.com.microsoft.teams: Operation not permitted\\n'
    };

    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);

    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === `ps ax | grep MacOS/MSTeams -m 1 | grep -v grep | awk '{ print $1 }'`) {
        return {};
      }
      if (opts === `rm -r "${homedir}/Library/Group Containers/UBF8T346G9.com.microsoft.teams"`) {
        return;
      }
      if (opts === `rm -r "${homedir}/Library/Containers/com.microsoft.teams2"`) {
        throw deleteError;
      }
      throw 'Invalid request';
    });
    await command.action(logger, { options: { force: true } } as any);
    assert(loggerLogSpy.calledWith('Deleting the folder failed. Please have a look at the following URL to delete the folders manually: https://answers.microsoft.com/en-us/msteams/forum/all/clearing-cache-on-microsoft-teams/35876f6b-eb1a-4b77-bed1-02ce3277091f'));
  });

  it('removes Teams cache from macOs platform without prompting using classic client', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: '' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        force: true,
        verbose: true,
        client: 'classic'
      }
    });
    assert(true);
  });

  it('removes teams cache when teams is currently not active using the classic client', async () => {
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
        verbose: true,
        client: 'classic'
      }
    });
    assert(true);
  });

  it('removes Teams cache from win32 platform without prompting using the classic client', async () => {
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
        verbose: true,
        client: 'classic'
      }
    });
    assert(true);
  });

  it('removes Teams cache from win32 platform without prompting using the new client', async () => {
    sinon.stub(process, 'platform').value('win32');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '', APPDATA: 'C:\\Users\\Administrator\\AppData\\Roaming', LOCALAPPDATA: 'C:\\Users\\Administrator\\AppData\\Local' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === 'wmic process where caption="ms-teams.exe" get ProcessId') {
        return { stdout: processOutput };
      }
      if (opts === 'rmdir /s /q "C:\\Users\\Administrator\\AppData\\Local\\Packages\\MSTeams_8wekyb3d8bbwe\\LocalCache\\Microsoft\\MSTeams"') {
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

  it('removes Teams cache from darwin platform with prompting using the classic client', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: 'pid' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        debug: true,
        client: 'classic'
      }
    });
    assert(true);
  });

  it('removes Teams cache from darwin platform with prompting', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(command, 'exec' as any).returns({ stdout: '1111' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        debug: true,
        client: 'new'
      }
    });
    assert(true);
  });

  it('removes teams cache when teams is currently not running on macOS', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(process, 'kill' as any).returns(null);
    sinon.stub(command, 'exec' as any).callsFake(async (opts) => {
      if (opts === `ps ax | grep MacOS/MSTeams -m 1 | grep -v grep | awk '{ print $1 }'`) {
        return {};
      }
      if (opts === `rm -r "${os.homedir()}/Library/Group Containers/UBF8T346G9.com.microsoft.teams"`) {
        return;
      }
      if (opts === `rm -r "${os.homedir()}/Library/Containers/com.microsoft.teams2"`) {
        return;
      }
      throw 'Invalid request';
    });
    sinon.stub(fs, 'existsSync').returns(true);

    await command.action(logger, {
      options: {
        force: true,
        verbose: true,
        client: 'new'
      }
    });
    assert(true);
  });


  it('aborts cache clearing when no cache folder is found using the classic client', async () => {
    sinon.stub(process, 'platform').value('darwin');
    sinon.stub(process, 'env').value({ 'CLIMICROSOFT365_ENV': '' });
    sinon.stub(fs, 'existsSync').returns(false);
    await command.action(logger, {
      options: {
        verbose: true,
        client: 'classic'
      }
    });
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

    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').resolves(false);

    await command.action(logger, { options: {} });
    assert(execStub.notCalled);
  });
});