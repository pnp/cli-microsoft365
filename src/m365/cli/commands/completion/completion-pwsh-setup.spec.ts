import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-pwsh-setup');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { autocomplete } from '../../../../autocomplete';
import * as chalk from 'chalk';

describe(commands.COMPLETION_PWSH_SETUP, () => {
  const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      fs.existsSync,
      fs.mkdirSync,
      fs.writeFileSync,
      fs.appendFileSync
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      autocomplete.generateShCompletion
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_PWSH_SETUP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('appends completion scripts to profile when profile file already exists', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, () => {
      try {
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('appends completion scripts to profile when profile file already exists (debug)', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: true, profile: profilePath } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWithExactly(chalk.green('DONE')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates profile file when it does not exist and appends the completion script to it', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, () => {
      try {
        assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates profile file when it does not exist and appends the completion script to it', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: true, profile: profilePath } }, () => {
      try {
        assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates profile path when it does not exist and appends the completion script to it', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(() => { });
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, () => {
      try {
        assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
        assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates profile path when it does not exist and appends the completion script to it (debug)', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(() => { });
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: true, profile: profilePath } }, () => {
      try {
        assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
        assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles exception when creating profile path', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const mkdirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(() => { throw error; });
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, (err?: any) => {
      try {
        assert(mkdirSyncStub.calledWith(path.dirname(profilePath), { recursive: true }), 'Profile path not created');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
        assert(writeFileSyncStub.notCalled, 'Profile file created');
        assert(appendFileSyncStub.notCalled, 'Completion script appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles exception when creating profile file', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'existsSync').callsFake((path) => path.toString().indexOf('.ps1') < 0);
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake((path) => { throw error; });
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, (err?: any) => {
      try {
        assert(writeFileSyncStub.calledWithExactly(profilePath, '', 'utf8'), 'Profile file not created');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
        assert(appendFileSyncStub.notCalled, 'Completion script appended');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles exception when appending completion script to the profile file', (done) => {
    const profilePath: string = '/Users/steve/.config/powershell/Microsoft.PowerShell_profile.ps1';
    const error: string = 'Unexpected error';
    sinon.stub(fs, 'writeFileSync').callsFake(() => { });
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const appendFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'appendFileSync').callsFake(() => { throw error; });

    cmdInstance.action({ options: { debug: false, profile: profilePath } }, (err?: any) => {
      try {
        assert(appendFileSyncStub.calledWithExactly(profilePath, os.EOL + completionScriptPath, 'utf8'), 'Completion script not appended');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
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