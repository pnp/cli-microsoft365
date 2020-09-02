import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./completion-pwsh-setup');
import * as assert from 'assert';
import Utils from '../../../../Utils';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { autocomplete } from '../../../../autocomplete';

describe(commands.COMPLETION_PWSH_SETUP, () => {
  const completionScriptPath: string = path.resolve(__dirname, '..', '..', '..', '..', '..', 'scripts', 'Register-CLIM365Completion.ps1');
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(autocomplete, 'generateShCompletion').callsFake(() => { });
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
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.COMPLETION_PWSH_SETUP), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert(cmdInstanceLogSpy.calledWithExactly(vorpal.chalk.green('DONE')));
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(error)), 'Invalid error returned');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the profile path is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation when the profile path specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { profile: 'profile.ps1' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.COMPLETION_PWSH_SETUP));
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