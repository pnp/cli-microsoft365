import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { fsUtil } from '../../../../utils/fsUtil';
import { packageManager } from '../../../../utils/packageManager';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
import { Manifest, Project, VsCode } from './project-model';
import { Finding, FindingToReport } from './report-model';
const command: Command = require('./project-upgrade');

describe(commands.PROJECT_UPGRADE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let trackEvent: any;
  let telemetryCommandName: any;
  let packagesDevExact: string[];
  let packagesDepExact: string[];
  let packagesDepUn: string[];
  let packagesDevUn: string[];
  let project141webPartNoLib: Project;
  const projectPath: string = './src/m365/spfx/commands/project/test-projects/spfx-141-webpart-nolib';

  before(() => {
    trackEvent = sinon.stub(telemetry, 'trackEvent').callsFake((commandName) => {
      telemetryCommandName = commandName;
    });
    project141webPartNoLib = (command as any).getProject(projectPath);
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
    telemetryCommandName = null;
    (command as any).allFindings = [];
    (command as any).packageManager = 'npm';
    (command as any).shell = 'bash';
    packagesDevExact = [];
    packagesDepExact = [];
    packagesDepUn = [];
    packagesDevUn = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      fs.statSync,
      fs.writeFileSync,
      fs.mkdirSync,
      fsUtil.readdirR
    ]);
  });

  after(() => {
    sinonUtil.restore([
      telemetry.trackEvent,
      pid.getProcessName
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PROJECT_UPGRADE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls telemetry', async () => {
    await assert.rejects(command.action(logger, { options: {} }));
    assert(trackEvent.called);
  });

  it('logs correct telemetry event', async () => {
    await assert.rejects(command.action(logger, { options: {} }));
    assert.strictEqual(telemetryCommandName, commands.PROJECT_UPGRADE);
  });

  it('shows error if the project path couldn\'t be determined', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`Couldn't find project root folder`, 1));
  });

  it('searches for package.json in the parent folder when it doesn\'t exist in the current folder', async () => {
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('package.json')) {
        return false;
      }
      else {
        return true;
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(`Couldn't find project root folder`, 1));
  });

  it('shows error if the specified spfx version is not supported by the CLI', async () => {
    await assert.rejects(command.action(logger, { options: { toVersion: '0.0.1' } } as any),
      new CommandError(`CLI for Microsoft 365 doesn't support upgrading SharePoint Framework projects to version 0.0.1. Supported versions are ${(command as any).supportedVersions.join(', ')}`, 2));
  });

  it('correctly handles the case when .yo-rc.json exists but doesn\'t contain spfx project info', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return `{}`;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any));
  });

  it('determines the current version from .yo-rc.json when available', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    const yoRcJson = `{
      "@microsoft/generator-sharepoint": {
        "version": "1.4.1",
        "libraryName": "spfx-141",
        "libraryId": "dd1a0a8d-e043-4ca0-b9a4-256e82a66177",
        "environment": "spo"
      }
    }`;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return yoRcJson;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.4.1' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('determines correct version from .yo-rc.json for an SP2019 project built using generator v1.10', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    const yoRcJson = `{
      "@microsoft/generator-sharepoint": {
        "isCreatingSolution": true,
        "environment": "onprem19",
        "version": "1.10.0",
        "libraryName": "spfx-1100-sp-2019",
        "libraryId": "04b9054d-025f-4e1a-9a85-732c57213b2f",
        "packageManager": "npm",
        "componentType": "webpart"
      }
    }`;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return yoRcJson;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.4.1' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('determines correct version from .yo-rc.json for an SP2016 project built using generator v1.10', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    const yoRcJson = `{
      "@microsoft/generator-sharepoint": {
        "isCreatingSolution": true,
        "environment": "onprem",
        "version": "1.10.0",
        "libraryName": "spfx-1100-sp-2016",
        "libraryId": "300833cb-9264-4b53-8179-2eaf105c1d41",
        "packageManager": "npm",
        "componentType": "webpart"
      }
    }`;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return yoRcJson;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.1.0' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.1.0');
  });

  it('tries to determine the current version from package.json if .yo-rc.json doesn\'t exist', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json')) {
        return `{
          "name": "spfx-141",
          "version": "0.0.1",
          "private": true,
          "engines": {
            "node": ">=0.10.0"
          },
          "scripts": {
            "build": "gulp bundle",
            "clean": "gulp clean",
            "test": "gulp test"
          },
          "dependencies": {
            "@microsoft/sp-core-library": "~1.4.1",
            "@microsoft/sp-webpart-base": "~1.4.1",
            "@microsoft/sp-lodash-subset": "~1.4.1",
            "@microsoft/sp-office-ui-fabric-core": "~1.4.1",
            "@types/webpack-env": ">=1.12.1 <1.14.0"
          },
          "devDependencies": {
            "@microsoft/sp-build-web": "~1.4.1",
            "@microsoft/sp-module-interfaces": "~1.4.1",
            "@microsoft/sp-webpart-workbench": "~1.4.1",
            "gulp": "~3.9.1",
            "@types/chai": ">=3.4.34 <3.6.0",
            "@types/mocha": ">=2.2.33 <2.6.0",
            "ajv": "~5.2.2"
          }
        }
        `;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.4.1' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('shows error if the project version couldn\'t be determined', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError(`Unable to determine the version of the current SharePoint Framework project`, 3));
  });

  it('determining project version doesn\'t fail if .yo-rc.json is empty', async () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return '';
      }
      else if (path.toString().endsWith('package.json')) {
        return `{
          "name": "spfx-141",
          "version": "0.0.1",
          "private": true,
          "engines": {
            "node": ">=0.10.0"
          },
          "scripts": {
            "build": "gulp bundle",
            "clean": "gulp clean",
            "test": "gulp test"
          },
          "dependencies": {
            "@microsoft/sp-core-library": "~1.4.1",
            "@microsoft/sp-webpart-base": "~1.4.1",
            "@microsoft/sp-lodash-subset": "~1.4.1",
            "@microsoft/sp-office-ui-fabric-core": "~1.4.1",
            "@types/webpack-env": ">=1.12.1 <1.14.0"
          },
          "devDependencies": {
            "@microsoft/sp-build-web": "~1.4.1",
            "@microsoft/sp-module-interfaces": "~1.4.1",
            "@microsoft/sp-webpart-workbench": "~1.4.1",
            "gulp": "~3.9.1",
            "@types/chai": ">=3.4.34 <3.6.0",
            "@types/mocha": ">=2.2.33 <2.6.0",
            "ajv": "~5.2.2"
          }
        }
        `;
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await command.action(logger, { options: { toVersion: '1.4.1' } } as any);
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
  });

  it('determining project version doesn\'t fail if package.json is empty', async () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('package.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    await assert.rejects(command.action(logger, { options: { toVersion: '1.4.1' } } as any));
    assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, undefined);
  });

  it('shows error if the current project version is not supported by the CLI', async () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '0.0.1');

    await assert.rejects(command.action(logger, { options: {} } as any),
      new CommandError(`CLI for Microsoft 365 doesn't support upgrading projects built using SharePoint Framework v0.0.1`, 4));
  });

  it('shows regular message if the current project version and the version to upgrade to are the same', async () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '1.5.0');

    await command.action(logger, { options: { toVersion: '1.5.0' } } as any);
    assert(log.indexOf(`Project doesn't need to be upgraded`) > -1, `Doesn't return info message`);
  });

  it('shows error if the current project version is higher than the version to upgrade to', async () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '1.5.0');

    await assert.rejects(command.action(logger, { options: { toVersion: '1.4.1' } } as any),
      new CommandError(`You cannot downgrade a project`, 5));
  });

  it('loads config.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.configJson), 'undefined');
  });

  it('loads copy-assets.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.copyAssetsJson), 'undefined');
  });

  it('loads deploy-azure-storage.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.deployAzureStorageJson), 'undefined');
  });

  it('loads package-solution.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.packageSolutionJson), 'undefined');
  });

  it('loads serve.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.serveJson), 'undefined');
  });

  it('loads tslint.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.tsLintJson), 'undefined');
  });

  it('loads write-manifests.json when available', () => {
    assert.notStrictEqual(typeof (project141webPartNoLib.writeManifestsJson), 'undefined');
  });

  it('doesn\'t fail if package.json not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('package.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.packageJson), 'undefined');
  });

  it('doesn\'t fail if tsconfig.json not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('tsconfig.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.tsConfigJson), 'undefined');
  });

  it('doesn\'t fail if config.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('config.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.configJson), 'undefined');
  });

  it('doesn\'t fail if copy-assets.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('copy-assets.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.copyAssetsJson), 'undefined');
  });

  it('doesn\'t fail if deploy-azure-storage.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('deploy-azure-storage.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.deployAzureStorageJson), 'undefined');
  });

  it('doesn\'t fail if package.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('package.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.packageJson), 'undefined');
  });

  it('doesn\'t fail if package-solution.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('package-solution.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.packageSolutionJson), 'undefined');
  });

  it('doesn\'t fail if serve.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('serve.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.serveJson), 'undefined');
  });

  it('doesn\'t fail if tslint.json is empty', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('tslint.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('tslint.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.tsLintJson), 'undefined');
  });

  it('doesn\'t fail if write-manifests.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('write-manifests.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.writeManifestsJson), 'undefined');
  });

  it('doesn\'t fail if .yo-rc.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof (project.yoRcJson), 'undefined');
  });

  it('doesn\'t fail if extensions.json is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('extensions.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof ((project.vsCode as VsCode).extensionsJson), 'undefined');
  });

  it('loads manifests when available', () => {
    assert.strictEqual((project141webPartNoLib.manifests as Manifest[]).length, 1);
  });

  it('doesn\'t fail if vscode settings are not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('settings.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof ((project.vsCode) as VsCode).settingsJson, 'undefined');
  });

  it('doesn\'t fail if vscode settings are empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('settings.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof ((project.vsCode as any).settingsJson), 'undefined');
  });

  it('doesn\'t fail if vscode launch info is empty', () => {
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, encoding) => {
      if (path.toString().endsWith('launch.json')) {
        return '';
      }
      else {
        return originalReadFileSync(path, encoding);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.strictEqual(typeof ((project.vsCode as any).launch), 'undefined');
  });

  //#region npm
  it(`doesn't return any dependencies from command npm for npm package manager`, () => {
    packageManager.mapPackageManagerCommand({
      command: 'npm', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'npm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dependency to be installed for npm i -SE for npm package manager`, () => {
    packageManager.mapPackageManagerCommand({
      command: 'npm i -SE package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'npm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 1, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dev dependency to be installed for npm i -DE for npm package manager`, () => {
    packageManager.mapPackageManagerCommand({
      command: 'npm i -DE package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'npm'
    });
    assert.strictEqual(packagesDevExact.length, 1, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dependency to uninstall for npm un -S for npm package manager`, () => {
    packageManager.mapPackageManagerCommand({
      command: 'npm un -S package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'npm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 1, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dev dependency to uninstall for npm un -D for npm package manager`, () => {
    packageManager.mapPackageManagerCommand({
      command: 'npm un -D package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'npm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 1, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns command to install dependency for 1 dep for npm package manager`, () => {
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: ['package'],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'npm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['npm i -SE package']));
  });

  it(`returns command to install dev dependency for 1 dev dep for npm package manager`, () => {
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: ['package'],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'npm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['npm i -DE package']));
  });

  it(`returns command to uninstall dependency for 1 dep for npm package manager`, () => {
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: ['package'],
      packagesDevUn: [],
      packageMgr: 'npm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['npm un -S package']));
  });

  it(`returns command to uninstall dev dependency for 1 dev dep for npm package manager`, () => {
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: ['package'],
      packageMgr: 'npm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['npm un -D package']));
  });
  //#endregion

  //#region pnpm
  it(`doesn't return any dependencies from command pnpm for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    packageManager.mapPackageManagerCommand({
      command: 'pnpm', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'pnpm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dependency to be installed for pnpm i -E for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    packageManager.mapPackageManagerCommand({
      command: 'pnpm i -E package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'pnpm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 1, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dev dependency to be installed for pnpm i -DE for npm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    packageManager.mapPackageManagerCommand({
      command: 'pnpm i -DE package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'pnpm'
    });
    assert.strictEqual(packagesDevExact.length, 1, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dev dependency to uninstall for pnpm un for npm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    packageManager.mapPackageManagerCommand({
      command: 'pnpm un package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'pnpm'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 1, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns command to install dependency for 1 dep for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: ['package'],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'pnpm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['pnpm i -E package']));
  });

  it(`returns command to install dev dependency for 1 dev dep for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: ['package'],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'pnpm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['pnpm i -DE package']));
  });

  it(`returns command to uninstall dependency for 1 dep for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: ['package'],
      packagesDevUn: [],
      packageMgr: 'pnpm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['pnpm un package']));
  });

  it(`returns command to uninstall dev dependency for 1 dev dep for pnpm package manager`, () => {
    (command as any).packageManager = 'pnpm';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: ['package'],
      packageMgr: 'pnpm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['pnpm un package']));
  });
  //#endregion

  //#region yarn
  it(`doesn't return any dependencies from command yarn for yarn package manager`, () => {
    (command as any).packageManager = 'yarn';
    packageManager.mapPackageManagerCommand({
      command: 'yarn', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'yarn'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dependency to be installed for yarn add -E for pnpm package manager`, () => {
    (command as any).packageManager = 'yarn';
    packageManager.mapPackageManagerCommand({
      command: 'yarn add -E package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'yarn'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 1, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dev dependency to be installed for yarn add -DE for npm package manager`, () => {
    (command as any).packageManager = 'yarn';
    packageManager.mapPackageManagerCommand({
      command: 'yarn add -DE package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'yarn'
    });
    assert.strictEqual(packagesDevExact.length, 1, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dev dependency to uninstall for yarn un for npm package manager`, () => {
    (command as any).packageManager = 'yarn';
    packageManager.mapPackageManagerCommand({
      command: 'yarn remove package', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn, packageMgr: 'yarn'
    });
    assert.strictEqual(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.strictEqual(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.strictEqual(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.strictEqual(packagesDevUn.length, 1, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns command to install dependency for 1 dep for yarn package manager`, () => {
    (command as any).packageManager = 'yarn';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: ['package'],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'yarn'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['yarn add -E package']));
  });

  it(`returns command to install dev dependency for 1 dev dep for yarn package manager`, () => {
    (command as any).packageManager = 'yarn';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: ['package'],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'yarn'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['yarn add -DE package']));
  });

  it(`returns command to uninstall dependency for 1 dep for yarn package manager`, () => {
    (command as any).packageManager = 'yarn';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: ['package'],
      packagesDevUn: [],
      packageMgr: 'yarn'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['yarn remove package']));
  });

  it(`returns command to uninstall dev dependency for 1 dev dep for yarn package manager`, () => {
    (command as any).packageManager = 'yarn';
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: ['package'],
      packageMgr: 'yarn'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify(['yarn remove package']));
  });
  //#endregion

  it(`returns no commands to run when no dependencies found`, () => {
    const commands: string[] = packageManager.reducePackageManagerCommand({
      packagesDepExact: [],
      packagesDevExact: [],
      packagesDepUn: [],
      packagesDevUn: [],
      packageMgr: 'npm'
    });
    assert.strictEqual(JSON.stringify(commands), JSON.stringify([]));
  });

  it('shows error when a upgrade rule failed', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-nolib'));
    (command as any).supportedVersions.splice(1, 0, '0');

    await assert.rejects(command.action(logger, { options: { toVersion: '1.0.1', output: 'json' } } as any), (err) => {
      (command as any).supportedVersions.splice(1, 1);
      return JSON.stringify(err).indexOf("Cannot find module './project-upgrade/upgrade-0'") > -1;
    });
  });

  //#region 1.0.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.0 project to 1.0.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.0.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 3);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.0.0 project to 1.0.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.0.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 3);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.0 project to 1.0.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.0.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 3);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.0.0 project to 1.0.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-100-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.0.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 3);
  });
  //#endregion

  //#region 1.0.1
  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.1 project to 1.0.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-nolib'));
    await command.action(logger, { options: { toVersion: '1.0.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 2);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.0.1 project to 1.0.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-react'));
    await command.action(logger, { options: { toVersion: '1.0.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 2);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.1 project to 1.0.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-ko'));
    await command.action(logger, { options: { toVersion: '1.0.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 5);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.0.1 project to 1.0.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-101-webpart-optionaldeps'));
    await command.action(logger, { options: { toVersion: '1.0.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 2);
  });
  //#endregion

  //#region 1.0.2
  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.2 project to 1.1.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.1.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.0.2 project to 1.1.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.1.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.2 project to 1.1.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.1.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.0.2 project to 1.1.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-102-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.1.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });
  //#endregion

  //#region 1.1.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.1.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 3);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.1.1', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 3);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.1.1', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 6);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.1.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 5);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.1.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 5);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.0 project to 1.1.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.1.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 5);
  });
  //#endregion

  //#region 1.1.1
  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.1.3', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 4);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.1.3', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 4);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.1.3', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 4);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.1.3', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 4);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.1.3', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 4);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.1 project to 1.1.3', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-111-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.1.3', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 4);
  });
  //#endregion

  //#region 1.1.3
  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 27);
  });

  it('e2e: shows correct number of findings for upgrading knockout web part 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.2.0', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 25);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.3 project to 1.2.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-113-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });
  //#endregion

  //#region 1.2.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.3.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.3.0', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.3.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.3.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.3.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.2.0 project to 1.3.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-120-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.3.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });
  //#endregion

  //#region 1.3.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.3.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.3.1', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.3.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.3.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.3.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.0 project to 1.3.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-130-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.3.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });
  //#endregion

  //#region 1.3.1
  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.3.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.3.2', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.3.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.3.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.3.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.1 project to 1.3.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-131-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.3.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 1);
  });
  //#endregion

  //#region 1.3.2
  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.3.4', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.3.4', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.3.4', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 18);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.3.4', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.3.4', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.2 project to 1.3.4', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-132-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.3.4', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });
  //#endregion

  //#region 1.3.4
  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.4 project to 1.4.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-134-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.4.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 25);
  });
  //#endregion

  //#region 1.4.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.4.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.4.1', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.4.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.4.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.4.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.4.0 project to 1.4.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-140-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.4.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });
  //#endregion

  //#region 1.4.1
  it('e2e: shows correct number of findings for upgrading no framework web part 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.5.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.5.0', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.5.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 33);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.5.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.5.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.4.1 project to 1.5.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.5.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 25);
  });
  //#endregion

  //#region 1.5.0
  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.5.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.5.1', debug: true, output: 'json' } } as any);
    const findings: Finding[] = log[3];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.5.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 18);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.5.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.5.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.5.0 project to 1.5.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-150-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.5.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });
  //#endregion

  //#region 1.5.1
  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project using MSGraphClient to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-nolib-graph'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 18);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project using AadHttpClient to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-nolib-aad'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 17);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.1 project using MSGraphClient to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 25);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.5.1 project to 1.6.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });
  //#endregion

  //#region 1.6.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 18);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 28);
  });

  it('e2e: suggests creating small teams app icon using a fixed name for upgrading react web part 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings[18].file, path.join('teams', 'tab20x20.png'));
  });

  it('e2e: suggests creating large teams app icon using a fixed name for upgrading react web part 1.6.0 project to 1.7.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.7.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings[19].file, path.join('teams', 'tab96x96.png'));
  });
  //#endregion

  //#region 1.7.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.7.0 project to 1.7.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-170-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.7.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });
  //#endregion

  //#region 1.7.1
  it('e2e: shows correct number of findings for upgrading application customizer 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 31);
  });

  it('e2e: suggests creating small teams app icon using a dynamic name for upgrading react web part 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings[20].file, path.join('teams', '7c4a6c24-2154-4dcc-9eb4-d64b8a2c5daa_outline.png'));
  });

  it('e2e: suggests creating large teams app icon using a dynamic name for upgrading react web part 1.7.1 project to 1.8.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings[21].file, path.join('teams', '7c4a6c24-2154-4dcc-9eb4-d64b8a2c5daa_color.png'));
  });
  //#endregion

  //#region 1.8.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.8.0 project to 1.8.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-180-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.8.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });
  //#endregion

  //#region 1.8.1
  it('e2e: shows correct number of findings for upgrading application customizer 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 17);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.8.1 project to 1.8.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-181-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });
  //#endregion

  //#region 1.8.2
  it('e2e: shows correct number of findings for upgrading application customizer 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 17);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.8.2 project to 1.9.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.9.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });
  //#endregion

  //#region 1.9.1
  it('e2e: shows correct number of findings for upgrading application customizer 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.9.1 project to 1.10.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-191-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.10.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });
  //#endregion

  //#region 1.10.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 20);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-ko'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 17);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 17);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.10.0 project to 1.11.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1100-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.11.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 27);
  });
  //#endregion

  //#region 1.11.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 28);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.11.0 project to 1.12.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1110-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.12.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 34);
  });
  //#endregion

  //#region 1.12.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.12.0 project to 1.12.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1120-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.12.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });
  //#endregion

  //#region 1.12.1
  it('e2e: shows correct number of findings for upgrading application customizer 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 18);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.12.1 project to 1.13.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1121-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.13.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 29);
  });
  //#endregion

  //#region 1.13.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 8);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 9);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.13.0 project to 1.13.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1130-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.13.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 20);
  });
  //#endregion

  //#region 1.13.1
  it('e2e: shows correct number of findings for upgrading application customizer 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.13.1 project to 1.14.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.14.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });
  //#endregion

  //#region 1.14.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 21);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 24);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 33);
  });

  it('e2e: shows correct number of findings for upgrading ace 1.14.0 project to 1.15.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1140-ace'));

    await command.action(logger, { options: { toVersion: '1.15.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 19);
  });
  //#endregion

  //#region 1.15.0
  it('e2e: shows correct number of findings for upgrading application customizer 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 14);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });

  it('e2e: shows correct number of findings for upgrading ace 1.15.0 project to 1.15.2', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1150-ace'));

    await command.action(logger, { options: { toVersion: '1.15.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });
  //#endregion

  //#region 1.15.2
  it('e2e: shows correct number of findings for upgrading ace 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-ace'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 20);
  });

  it('e2e: shows correct number of findings for upgrading form customizer react 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-formcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 22);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 15);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 16);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.15.2 project to 1.16.0', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1152-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.16.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 26);
  });
  //#endregion

  //#region 1.16.0
  it('e2e: shows correct number of findings for upgrading ace 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-ace'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 10);
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-applicationcustomizer'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-fieldcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 11);
  });

  it('e2e: shows correct number of findings for upgrading form customizer react 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-formcustomizer-react'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-listviewcommandset'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 12);
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-webpart-nolib'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 13);
  });

  it('e2e: shows correct number of findings for upgrading web part with optional dependencies 1.16.0 project to 1.16.1', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1160-webpart-optionaldeps'));

    await command.action(logger, { options: { toVersion: '1.16.1', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });
  //#endregion

  //#region superseded rules
  it('ignores superseded findings (1.1.0 > 1.2.0)', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-110-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.2.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 23);
  });

  it('ignores superseded findings (1.6.0 > 1.8.0)', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-160-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 33);
  });

  it('ignores superseded findings (1.7.1 > 1.8.2)', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-171-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.8.2', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 27);
  });

  it('ignores superseded findings (1.4.1 > 1.6.0)', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-141-webpart-react'));

    await command.action(logger, { options: { toVersion: '1.6.0', output: 'json' } } as any);
    const findings: FindingToReport[] = log[0];
    assert.strictEqual(findings.length, 32);
  });
  //#endregion

  it('shows all information with output format json', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-fieldcustomizer-react'));

    await command.action(logger, { options: { output: 'json' } } as any);
    assert(JSON.stringify(log[0]).indexOf('"resolution":') > -1);
  });

  it('upgrades project to the latest preview version using the preview option', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-1131-webpart-nolib'));

    await command.action(logger, { options: { output: 'text' } } as any);
    assert(log[0].indexOf('1.15.2') > -1);
  });

  it('returns markdown report with output format md', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph'));

    await command.action(logger, { options: { output: 'md', toVersion: '1.6.0' } } as any);
    assert(log[0].indexOf('## Findings') > -1);
  });

  it('overrides base md formatting', async () => {
    const expected = [
      {
        'prop1': 'value1'
      },
      {
        'prop2': 'value2'
      }
    ];
    const actual = command.getMdOutput(expected, command, { options: { output: 'md' } } as any);
    assert.deepStrictEqual(actual, expected);
  });

  it('returns json report with output format default', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph'));

    await command.action(logger, { options: { toVersion: '1.6.0' } } as any);
    assert(JSON.stringify(log[0]).indexOf('"resolution":') > -1);
  });

  it('returns text report with output format text', async () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph'));

    await command.action(logger, { options: { output: 'text', toVersion: '1.6.0' } } as any);
    assert(log[0].indexOf('Execute in ') > -1);
  });

  it('writes CodeTour upgrade report to .tours folder when in tour output mode. Creates the folder when it does not exist', async () => {
    const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph';
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return false;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');

    await command.action(logger, { options: { output: 'tour', toVersion: '1.6.0' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.tours/upgrade.tour')), 'Tour file not created');
    assert(mkDirSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.tours')), '.tours folder not created');
  });

  it('writes CodeTour upgrade report to .tours folder when in tour output mode. Does not create the folder when it already exists', async () => {
    const projectPath: string = 'src/m365/spfx/commands/project/test-projects/spfx-151-webpart-react-graph';
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), projectPath));
    const writeFileSyncStub: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    const existsSyncOriginal = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake(path => {
      if (path.toString().indexOf('.tours') > -1) {
        return true;
      }

      return existsSyncOriginal(path);
    });
    const mkDirSyncStub: sinon.SinonStub = sinon.stub(fs, 'mkdirSync').callsFake(_ => '');

    await command.action(logger, { options: { output: 'tour', toVersion: '1.6.0' } } as any);
    assert(writeFileSyncStub.calledWith(path.join(process.cwd(), projectPath, '/.tours/upgrade.tour')), 'Tour file not created');
    assert(mkDirSyncStub.notCalled, '.tours folder created');
  });

  it('passes validation when package manager not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when npm package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'npm' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when pnpm package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'pnpm' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when yarn package manager specified', async () => {
    const actual = await command.validate({ options: { packageManager: 'yarn' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when shell not specified', async () => {
    const actual = await command.validate({ options: {} }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when unsupported shell specified', async () => {
    const actual = await command.validate({ options: { shell: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when bash shell specified', async () => {
    const actual = await command.validate({ options: { shell: 'bash' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when powershell shell specified', async () => {
    const actual = await command.validate({ options: { shell: 'powershell' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when cmd shell specified', async () => {
    const actual = await command.validate({ options: { shell: 'cmd' } }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
