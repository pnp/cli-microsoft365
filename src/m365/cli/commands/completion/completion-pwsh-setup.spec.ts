import assert from 'assert';
import fs from 'fs';
import os from 'os';
import path from 'path';
import sinon from 'sinon';
import url from 'url';
import { CommandError } from '../../../../Command.js';
import { autocomplete } from '../../../../autocomplete.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './completion-pwsh-setup.js';

const __dirname = url.fileURLToPath(new URL('.', import.meta.url));

describe(commands.COMPLETION_PWSH_SETUP, () => {
  const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
  let log: string[];
  let logger: Logger;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.mkdirSync,
      fs.writeFileSync,
      fs.appendFileSync
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_PWSH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('appends completion scripts to profile when profile file already exists', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { profile: profilePath } });
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'));
  });

  it('appends completion scripts to profile when profile file already exists (debug)', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { debug: true, profile: profilePath } });
    assert(loggerLogToStderrSpy.called);
  });

  it('creates profile file when it does not exist and appends the completion script to it', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { profile: profilePath } });
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
  });

  it('creates profile file when it does not exist and appends the completion script to it (debug)', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { debug: true, profile: profilePath } });
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
  });

  it('creates profile path when it does not exist and appends the completion script to it', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { profile: profilePath } });
    assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
  });

  it('creates profile path when it does not exist and appends the completion script to it (debug)', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await command.action(logger, { options: { debug: true, profile: profilePath } });
    assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
  });

  it('handles exception when creating profile path', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(() => { throw error; });
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
    assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
    assert(writeFileSyncStub.notCalled, 'Profile file created');
    assert(appendFileSyncStub.notCalled, 'Completion script appended');
  });

  it('handles exception when creating profile file', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { throw error; });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
    assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
    assert(appendFileSyncStub.notCalled, 'Completion script appended');
  });

  it('handles exception when appending completion script to the profile file', async () => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { throw error; });

    await assert.rejects(command.action(logger, { options: { profile: profilePath } } as any), new CommandError(error));
    assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
  });
});
