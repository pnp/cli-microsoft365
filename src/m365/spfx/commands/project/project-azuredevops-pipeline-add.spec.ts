import assert from 'assert';
import fs from 'fs';
import path from 'path';
import sinon from 'sinon';
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
import command from './project-azuredevops-pipeline-add.js';

describe(commands.PROJECT_AZUREDEVOPS_PIPELINE_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const projectPath: string = path.resolve('/fake/path/to/test-project');

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(spfx, 'getHighestNodeVersion').returns('22.0.x');
    sinon.stub(session, 'getId').callsFake(() => '');
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
      if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((fakePath, options) => {
      if (fakePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
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

    await command.action(logger, { options: { name: 'test', branchName: 'dev', skipFeatureDeployment: true, loginMethod: 'user', scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' } } as any);
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.azuredevops', 'pipelines', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('fails validation if loginMethod is not valid type', async () => {
    const actual = await command.validate({ options: { loginMethod: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is not valid type', async () => {
    const actual = await command.validate({ options: { scope: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if scope is sitecollection but the siteUrl was not defined', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if siteUrl is not valid', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all required properties are provided', async () => {
    const actual = await command.validate({ options: { scope: 'sitecollection', siteUrl: 'https://contoso.sharepoint.com/sites/project' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('creates a default workflow (debug)', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);
    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
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

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: { debug: true } } as any);
    assert(writeFileSyncStub.calledWith(path.resolve(path.join(projectPath, '.azuredevops', 'pipelines', 'deploy-spfx-solution.yml'))), 'workflow file not created');
  });

  it('handles error with unknown minor version of SPFx when missing minor version', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on the version property in the .yo-rc.json file.'));
  });

  it('handles error with not found node version', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('99.99.99');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Could not find Node version for version '99.99.99' of SharePoint Framework.`));
  });

  it('handles unexpected error', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('writeFileSync failed'));
  });

  it('handles unexpected non-error value', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(projectPath);

    sinon.stub(fs, 'readFileSync').callsFake((filePath, options) => {
      if (filePath.toString() === path.join(projectPath, 'package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      throw `Invalid path: ${filePath}`;
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString() === path.join(projectPath, '.azuredevops')) {
        return true;
      }
      else if (fakePath.toString() === path.join(projectPath, '.azuredevops', 'pipelines')) {
        return true;
      }

      throw `Invalid path: ${fakePath}`;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').callsFake(() => {
      throw 'string failure';
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('string failure'));
  });
});
