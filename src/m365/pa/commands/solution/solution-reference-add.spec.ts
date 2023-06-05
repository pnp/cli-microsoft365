import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import CdsProjectMutator from '../../cds-project-mutator';
import commands from '../../commands';
const command: Command = require('./solution-reference-add');

describe(commands.SOLUTION_REFERENCE_ADD, () => {
  let cli: Cli;
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    cli = Cli.getInstance();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    sinon.stub(cli, 'getSettingWithDefaultValue').callsFake(((settingName, defaultValue) => defaultValue));
  });

  afterEach(() => {
    sinonUtil.restore([
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      path.relative,
      fs.readdirSync,
      CdsProjectMutator.prototype.addProjectReference,
      cli.getSettingWithDefaultValue
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.SOLUTION_REFERENCE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('supports specifying projectPath', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--projectPath') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation when no *.cdsproj exists in the current directory', async () => {
    sinon.stub(fs, 'readdirSync').returns([]);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when more than one *.cdsproj exists in the current directory', async () => {
    sinon.stub(fs, 'readdirSync').callsFake(() => ['file1.cdsproj', 'file2.cdsproj'] as any);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the path option isn\'t specified', async () => {
    sinon.stub(fs, 'readdirSync').returns(['file1.cdsproj'] as any);

    const actual = await command.validate({ options: {} }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified path option doesn\'t exist', async () => {
    sinon.stub(fs, 'readdirSync').returns(['file1.cdsproj'] as any);
    sinon.stub(fs, 'existsSync').returns(false);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified path contains no *.pcfproj or *.csproj file', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return [];
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified path contains two *.pcfproj files', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj', 'file2.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified path contains two *.csproj files', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.csproj', 'file2.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when the specified path contains both *.pcfproj and *.csproj files', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj', 'file2.csproj', 'file3.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.pcfproj files', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['cdsfile1.cdsproj'] as any;
      }
      return ['pcffile1.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.csproj files', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['cdsfile1.cdsproj'] as any;
      }
      return ['csfile1.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.pcfproj file with the same name', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.pcfproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when current directory contains exactly one *.cdsproj file and the specified path contains exactly one *.csproj file with the same name', async () => {
    sinon.stub(fs, 'readdirSync').callsFake((path) => {
      if (path === process.cwd()) {
        return ['file1.cdsproj'] as any;
      }
      return ['file1.csproj'] as any;
    });
    sinon.stub(fs, 'existsSync').returns(true);

    const actual = await command.validate({ options: { projectPath: 'path/to/project' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('Creates an instance of CdsProjectMutator, adds project reference and saves updated file', async () => {
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
    const pathRelative = sinon.stub(path, 'relative').returns(pathToPcfProject);
    const fsReadFileSync = sinon.stub(fs, 'readFileSync').returns('<abc></abc>');
    const addProjectReferenceStub = sinon.stub(CdsProjectMutator.prototype, 'addProjectReference').returns();
    const fsWriteFileSync = sinon.stub(fs, 'writeFileSync').returns();

    await command.action(logger, { options: { projectPath: pathToDirectory } });
    assert(pathRelative.calledWith(process.cwd(), pathToPcfProject));
    assert(fsReadFileSync.calledWith(pathToCdsProject, 'utf8'));
    assert(addProjectReferenceStub.calledWith(pathToPcfProject));
    assert(fsWriteFileSync.calledWith(pathToCdsProject, sinon.match.any));
  });

  it('throws error if path does not exist', async () => {
    const pathToDirectory = '../path/to/projectDirectory';
    sinon.stub(fs, 'readdirSync').throws(new Error('An error has occured'));

    await assert.rejects(command.action(logger, { options: { projectPath: pathToDirectory } }), new Error('An error has occured'));
  });


  it('supports verbose mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach((o) => {
      if (o.option === '--verbose') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
