import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./solution-reference-add');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import CdsProjectMutator from '../../cds-project-mutator';

describe(commands.SOLUTION_REFERENCE_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
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
    telemetry = null;
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      path.relative,
      fs.readdirSync,
      CdsProjectMutator.prototype.addProjectReference
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.SOLUTION_REFERENCE_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('calls telemetry', () => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      assert(trackEvent.called);
    });
  });

  it('logs correct telemetry event', () => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      assert.equal(telemetry.name, commands.SOLUTION_REFERENCE_ADD);
    });
  });

  it('supports specifying path', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--path') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation when no *.cdsproj exists in the current directory', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => []);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when more than one *.cdsproj exists in the current directory', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.cdsproj', 'file2.cdsproj'] as any);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the path option isn\'t specified', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.cdsproj'] as any);

    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path option doesn\'t exist', () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.cdsproj'] as any);
    sinon.stub(fs, 'existsSync').callsFake(() => false);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path contains no *.pcfproj or *.csproj file', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return [];
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path contains no *.pcfproj or *.csproj file', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return [];
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path contains two *.pcfproj files', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj', 'file2.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path contains two *.csproj files', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.csproj', 'file2.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when the specified path contains both *.pcfproj and *.csproj files', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj', 'file2.csproj', 'file3.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.pcfproj files', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['cdsfile1.cdsproj'] as any;
      }
      return ['pcffile1.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.equal(actual, true);
  });

  it('passes validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.csproj files', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['cdsfile1.cdsproj'] as any;
      }
      return ['csfile1.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.equal(actual, true);
  });

  it('fails validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.pcfproj file with the same name', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('fails validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.csproj file with the same name', () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').callsFake(() => true);

    const actual = (command.validate() as CommandValidate)({ options: { path: 'path/to/project' } });
    assert.notEqual(actual, true);
  });

  it('Creates an instance of CdsProjectMutator, adds project reference and saves updated file', () => {
    const pathToDirectory = '../path/to/projectDirectory';
    const pcfProjectFile = 'project.pcfproj';
    const pathToPcfProject = path.join(pathToDirectory, pcfProjectFile);
    const cdsProjectFile = 'cdsproject.cdsproj';
    const pathToCdsProject = path.join(process.cwd(), cdsProjectFile);
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === pathToDirectory) {
        return [pcfProjectFile] as any;
      }
      else if (path === process.cwd()) {
        return [cdsProjectFile] as any;
      }
      return [];
    });
    const pathRelative = sinon.stub(path, 'relative').callsFake((from, to) => {
      return pathToPcfProject;
    });
    const fsReadFileSync = sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => '<abc></abc>');
    const addProjectReferenceStub = sinon.stub(CdsProjectMutator.prototype, 'addProjectReference').callsFake((path) => { });
    const fsWriteFileSync = sinon.stub(fs, 'writeFileSync').callsFake((path, contents) => { });

    cmdInstance.action({ options: { path: pathToDirectory } }, () => {
      assert(pathRelative.calledWith(process.cwd(), pathToPcfProject));
      assert(fsReadFileSync.calledWith(pathToCdsProject, 'utf8'));
      assert(addProjectReferenceStub.calledWith(pathToPcfProject));
      assert(fsWriteFileSync.calledWith(pathToCdsProject, sinon.match.any));
    });
  });

  it('supports verbose mode', () => {
    const options = command.options() as CommandOption[];
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
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
    assert(find.calledWith(commands.SOLUTION_REFERENCE_ADD));
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