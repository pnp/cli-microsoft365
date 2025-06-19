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
import command from './project-github-workflow-add.js';

describe(commands.PROJECT_GITHUB_WORKFLOW_ADD, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  const projectPath: string = 'test-project';

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
    assert.strictEqual(command.name, commands.PROJECT_GITHUB_WORKFLOW_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
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
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: { debug: true } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.github', 'workflows', 'deploy-spfx-solution.yml')), 'workflow file not created');
  });

  it('creates a default workflow with specifying options', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(fs, 'mkdirSync').callsFake((path, options) => {
      if (path.toString().endsWith('.github') && (options as fs.MakeDirectoryOptions).recursive) {
        return `${projectPath}/.github`;
      }

      return '';
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').resolves({});

    await command.action(logger, { options: { name: 'test', branchName: 'dev', manuallyTrigger: true, skipFeatureDeployment: true, loginMethod: 'user', scope: 'sitecollection' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.github', 'workflows', 'deploy-spfx-solution.yml')), 'workflow file not created');
  });

  it('handles error with unknown version of SPFx', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(command as any, 'getProjectVersion').returns(undefined);

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError('Unable to determine the version of the current SharePoint Framework project. Could not find the correct version based on @microsoft/generator-sharepoint property in the .yo-rc.json file.'));

  });

  it('handles error with not found node version', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('99.99.99');

    sinon.stub(fs, 'writeFileSync').throws(new Error('writeFileSync failed'));

    await assert.rejects(command.action(logger, { options: {} }),
      new CommandError(`Could not find Node version for 99.99.99 of SharePoint Framework`));
  });

  it('handles unexpected error', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').callsFake(() => {
      throw new Error('writeFileSync failed');
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('writeFileSync failed'));
  });

  it('handles unexpected non-error value', async () => {
    sinon.stub(command as any, 'getProjectRoot').returns(path.join(process.cwd(), projectPath));

    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && options === 'utf-8') {
        return '{"name": "test"}';
      }

      return '';
    });

    sinon.stub(fs, 'existsSync').callsFake((fakePath) => {
      if (fakePath.toString().endsWith('.github')) {
        return true;
      }
      else if (fakePath.toString().endsWith('workflows')) {
        return true;
      }

      return false;
    });

    sinon.stub(command as any, 'getProjectVersion').returns('1.21.1');

    sinon.stub(fs, 'writeFileSync').callsFake(() => {
      throw 'string failure';
    });

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError('string failure'));
  });
});
