import assert from 'assert';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
import yaml from 'yaml';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { spfx } from '../../../../utils/spfx.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command, { options } from './project-github-workflow-add.js';
import { GitHubWorkflow } from './project-github-workflow-model.js';

describe(commands.PROJECT_GITHUB_WORKFLOW_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const projectPath: string = path.resolve('/test-project');

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spfx, 'getHighestNodeVersion').returns('22.0.x');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.PROJECT_GITHUB_WORKFLOW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if loginMethod is not valid type', () => {
    const actual = commandOptionsSchema.safeParse({ loginMethod: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if scope is not valid type', () => {
    const actual = commandOptionsSchema.safeParse({ scope: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if scope is sitecollection but the siteUrl was not defined', () => {
    const actual = commandOptionsSchema.safeParse({ scope: 'sitecollection' });
    assert.strictEqual(actual.success, false);
  });

  it('fails validation if siteUrl is not valid', () => {
    const actual = commandOptionsSchema.safeParse({ scope: 'sitecollection', siteUrl: 'abc' });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation if all required properties are provided', () => {
    const actual = commandOptionsSchema.safeParse({ scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' });
    assert.strictEqual(actual.success, true);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(null);

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('creates a default workflow (debug)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: commandOptionsSchema.parse({ debug: true }) });
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.github', 'workflows', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('creates a default workflow with specifying options', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'mkdirSync').callsFake((fakePath, options) => {
      if (fakePath.toString() === path.join(projectPath, '.github') && (options as fs.MakeDirectoryOptions).recursive) {
        return path.join(projectPath, '.github');
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'test', branchName: 'dev', manuallyTrigger: true, skipFeatureDeployment: true, loginMethod: 'user', scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/test' }) });
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.github', 'workflows', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('uses npm run build and package for newer version of SPFx', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.yo-rc.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, '.yo-rc.json') && options === 'utf-8') {
        return '{"@microsoft/generator-sharepoint": {"version": "1.22.0"}}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').returns();

    await command.action(logger, { options: commandOptionsSchema.parse({}) });

    assert(writeFileSyncStub.calledOnce, 'writeFileSync not called or called multiple times.');
    const writtenWorkflow: GitHubWorkflow = yaml.parse(writeFileSyncStub.args[0][1] as string);
    const buildStep = writtenWorkflow.jobs['build-and-deploy'].steps.find(step => step.name === 'Build & Package');
    assert.strictEqual(buildStep?.run, 'npm run build', 'Build & Package step does not run npm run build');
  });

  it('handles error with unknown version of SPFx', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns(undefined);

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.'));

  });

  it('handles error with not found node version', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('99.99.99');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }), new CommandError("Could not find Node version for version '99.99.99' of SharePoint Framework."));
  });

  it('handles unexpected error', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (filePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.github')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.github', 'workflows')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').throws(
      new Error('writeFileSync failed')
    );

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('writeFileSync failed'));
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      name: 'test',
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('passes validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.strictEqual(actual.success, true);
  });
});