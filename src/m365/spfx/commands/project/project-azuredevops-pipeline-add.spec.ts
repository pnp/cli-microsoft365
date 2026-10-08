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
import command, { options } from './project-azuredevops-pipeline-add.js';
import { AzureDevOpsPipeline } from './project-azuredevops-pipeline-model.js';

describe(commands.PROJECT_AZUREDEVOPS_PIPELINE_ADD, () => {
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
    assert.strictEqual(command.name, commands.PROJECT_AZUREDEVOPS_PIPELINE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a default workflow with specifying options', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((fakePath, options) => {
      if (fakePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json') && options === 'utf-8') {
        return '{"paths": {"zippedPackage": "solution/test.sppkg"}}';
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(fs, 'mkdirSync').callsFake((fakePath, options) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops') && (options as fs.MakeDirectoryOptions).recursive) {
        return path.join(projectPath, '.azuredevops');
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.16.0');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: commandOptionsSchema.parse({ name: 'test', branchName: 'dev', skipFeatureDeployment: true, loginMethod: 'user', scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' }) });
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.azuredevops', 'pipelines', 'deploy-spfx-solution.yml'))), 'workflow file not created');
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
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
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
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.azuredevops', 'pipelines', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('creates a pipeline with npm run build for SPFx version that requires heft', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);
    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
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

    sinon.stub(command as any, 'getProjectVersion').returns('1.22.0');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });

    await command.action(logger, { options: commandOptionsSchema.parse({}) });

    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.azuredevops', 'pipelines', 'deploy-spfx-solution.yml'))), 'workflow file not created');
    const writtenPipeline: AzureDevOpsPipeline = yaml.parse(writeFileSyncStub.args[0][1] as string);
    const steps = writtenPipeline.stages[0].jobs[0].steps;
    const buildStep = steps.find(step => step.displayName === 'Build and package');
    assert.strictEqual(buildStep?.inputs?.customCommand, 'run build', 'Build and package step does not run npm run build');
    assert.strictEqual(steps.some(step => step.task === 'Gulp@0'), false, 'Gulp steps should not be added');
  });

  it('handles error with unknown minor version of SPFx when missing minor version', async () => {
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
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns(undefined);

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.'));
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
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('99.99.99');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: commandOptionsSchema.parse({}) }),
      new CommandError(`Could not find Node version for version '99.99.99' of SharePoint Framework.`));
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
      if (fakePath.toString() === path.join(projectPath, 'package.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, 'config', 'package-solution.json')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

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