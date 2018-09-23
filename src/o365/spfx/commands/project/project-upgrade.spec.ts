import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./project-upgrade');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import { Utils as Utils1 } from './project-upgrade/';
import { Project, Manifest, VsCode } from './project-upgrade/model';
import { Finding } from './project-upgrade/Finding';

describe(commands.PROJECT_UPGRADE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  let packagesDevExact: string[];
  let packagesDepExact: string[];
  let packagesDepUn: string[];
  let packagesDevUn: string[];

  before(() => {
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    telemetry = null;
    (command as any).allFindings = [];
    packagesDevExact = [];
    packagesDepExact = [];
    packagesDepUn = [];
    packagesDevUn = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      Utils1.getAllFiles
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PROJECT_UPGRADE), true);
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
      assert.equal(telemetry.name, commands.PROJECT_UPGRADE);
    });
  });

  it('shows error if the project path couldn\'t be determined', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`)));
    });
  });

  it('searches for package.json in the parent folder when it doesn\'t exist in the current folder', () => {
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('package.json')) {
        return false;
      }
      else {
        return true;
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`)));
    });
  });

  it('shows error if the specified spfx version is not supported by the CLI', () => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '0.0.1' } }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Office 365 CLI doesn't support upgrading SharePoint Framework projects to version 0.0.1. Supported versions are ${(command as any).supportedVersions.join(', ')}`)));
    });
  });

  it('correctly handles the case when .yo-rc.json exists but doesn\'t contain spfx project info', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return `{}`;
      }
      else {
        return originalReadFileSync(path);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert(true);
    });
  });

  it('determines the current version from .yo-rc.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return `{
          "@microsoft/generator-sharepoint": {
            "version": "1.4.1",
            "libraryName": "spfx-141",
            "libraryId": "dd1a0a8d-e043-4ca0-b9a4-256e82a66177",
            "environment": "spo"
          }
        }`;
      }
      else {
        return originalReadFileSync(path);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert(true);
    });
  });

  it('tries to determine the current version from package.json if .yo-rc.json doesn\'t exist', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('package.json')) {
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
        return originalReadFileSync(path);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert(true);
    });
  });

  it('shows error if the project version couldn\'t be determined', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('.yo-rc.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Unable to determine the version of the current SharePoint Framework project`)));
    });
  });

  it('shows error if the current project version is not supported by the CLI', () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '0.0.1');

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Office 365 CLI doesn't support upgrading projects build on SharePoint Framework v0.0.1`)));
    });
  });

  it('shows error if the current project version and the version to upgrade to are the same', () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '1.5.0');

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0' } }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Project doesn't need to be upgraded`)));
    });
  });

  it('shows error if the current project version is higher than the version to upgrade to', () => {
    sinon.stub(command as any, 'getProjectVersion').callsFake(_ => '1.5.0');

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1' } }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`You cannot downgrade a project`)));
    });
  });

  it('loads config.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('config.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('config.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.configJson), 'undefined');
  });

  it('loads copy-assets.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('copy-assets.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('copy-assets.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.copyAssetsJson), 'undefined');
  });

  it('loads deploy-azure-storage.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('deploy-azure-storage.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('deploy-azure-storage.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.deployAzureStorageJson), 'undefined');
  });

  it('loads package-solution.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('package-solution.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('package-solution.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.packageSolutionJson), 'undefined');
  });

  it('loads serve.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('serve.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('serve.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.serveJson), 'undefined');
  });

  it('loads tslint.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('tslint.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('tslint.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.tsLintJson), 'undefined');
  });

  it('loads write-manifests.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('write-manifests.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('write-manifests.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.notEqual(typeof (project.writeManifestsJson), 'undefined');
  });

  it('doesn\'t fail if package.json not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('package.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.equal(typeof (project.packageJson), 'undefined');
  });

  it('doesn\'t fail if tsconfig.json not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('tsconfig.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.equal(typeof (project.tsConfigJson), 'undefined');
  });

  it('loads manifests when available', () => {
    sinon.stub(Utils1, 'getAllFiles').callsFake(_ => [
      '/usr/tmp/HelloWorldWebPart.ts',
      '/usr/tmp/HelloWorldWebPart.manifest.json'
    ]);
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('.manifest.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.equal((project.manifests as Manifest[]).length, 1);
  });

  it('doesn\'t fail if vscode settings are not available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path: string) => {
      if (path.endsWith('settings.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject('./');
    assert.equal(typeof ((project.vsCode) as VsCode).settingsJson, 'undefined');
  });

  it(`doesn't return any dependencies from npm`, () => {
    (command as any).mapNpmCommand('npm', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dependency to be installed for npm i -SE`, () => {
    (command as any).mapNpmCommand('npm i package -SE', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 1, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 exact dev dependency to be installed for npm i -DE`, () => {
    (command as any).mapNpmCommand('npm i package -DE', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 1, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`doesn't return any dependencies for npm i -S`, () => {
    (command as any).mapNpmCommand('npm i package -S', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`doesn't return any dependencies for npm i -D`, () => {
    (command as any).mapNpmCommand('npm i package -D', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`doesn't return any dependencies for npm un -SE`, () => {
    (command as any).mapNpmCommand('npm un package -SE', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`doesn't return any dependencies for npm un -DE`, () => {
    (command as any).mapNpmCommand('npm un package -DE', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dependency to uninstall for npm un -S`, () => {
    (command as any).mapNpmCommand('npm un package -S', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 1, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 0, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns 1 dev dependency to uninstall for npm un -D`, () => {
    (command as any).mapNpmCommand('npm un package -D', packagesDevExact, packagesDepExact, packagesDepUn, packagesDevUn);
    assert.equal(packagesDevExact.length, 0, 'Incorrect number of deps to install');
    assert.equal(packagesDepExact.length, 0, 'Incorrect number of dev deps to install');
    assert.equal(packagesDepUn.length, 0, 'Incorrect number of deps to uninstall');
    assert.equal(packagesDevUn.length, 1, 'Incorrect number of dev deps to uninstall');
  });

  it(`returns no commands to run when no dependencies found`, () => {
    const commands: string[] = (command as any).reduceNpmCommand([], [], [], []);
    assert.equal(JSON.stringify(commands), JSON.stringify([]));
  });

  it(`returns command to install dependency for 1 dep`, () => {
    const commands: string[] = (command as any).reduceNpmCommand(['package'], [], [], []);
    assert.equal(JSON.stringify(commands), JSON.stringify(['npm i package -SE']));
  });

  it(`returns command to install dev dependency for 1 dev dep`, () => {
    const commands: string[] = (command as any).reduceNpmCommand([], ['package'], [], []);
    assert.equal(JSON.stringify(commands), JSON.stringify(['npm i package -DE']));
  });

  it(`returns command to uninstall dependency for 1 dep`, () => {
    const commands: string[] = (command as any).reduceNpmCommand([], [], ['package'], []);
    assert.equal(JSON.stringify(commands), JSON.stringify(['npm un package -S']));
  });

  it(`returns command to uninstall dev dependency for 1 dev dep`, () => {
    const commands: string[] = (command as any).reduceNpmCommand([], [], [], ['package']);
    assert.equal(JSON.stringify(commands), JSON.stringify(['npm un package -D']));
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.0 project to 1.0.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-100-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 3);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.0.0 project to 1.0.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-100-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 3);
    });
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.0 project to 1.0.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-100-webpart-ko'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 3);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.1 project to 1.0.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-101-webpart-nolib'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 2);
    });
  });
 
  it('e2e: shows correct number of findings for upgrading react web part 1.0.1 project to 1.0.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-101-webpart-react'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 2);
    });
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.1 project to 1.0.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-101-webpart-ko'));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.0.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.0.2 project to 1.1.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.0.2 project to 1.1.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 23);
    });
  });

  it('e2e: shows correct number of findings for upgrading ko web part 1.0.2 project to 1.1.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-102-webpart-ko'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.0 project to 1.1.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-110-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 3);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.0 project to 1.1.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-110-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.1', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 3);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.0 project to 1.1.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-110-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.0 project to 1.1.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-110-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.0 project to 1.1.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-110-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.1 project to 1.1.3', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-111-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.3', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.1 project to 1.1.3', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-111-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.3', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 5);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.1 project to 1.1.3', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-111-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.3', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 4);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.1 project to 1.1.3', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-111-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.3', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 4);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.1 project to 1.1.3', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-111-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.1.3', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 4);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 21);
    });
  });

  it('e2e: shows correct number of findings for upgrading knockout web part 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-webpart-ko'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 21);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 22);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 23);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 25);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.1.3 project to 1.2.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-113-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.2.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 24);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.2.0 project to 1.3.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-120-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.2.0 project to 1.3.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-120-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.0', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.2.0 project to 1.3.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-120-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 10);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.2.0 project to 1.3.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-120-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 10);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.2.0 project to 1.3.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-120-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 10);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.0 project to 1.3.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-130-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.0 project to 1.3.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-130-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.1', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.0 project to 1.3.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-130-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.0 project to 1.3.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-130-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.0 project to 1.3.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-130-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.1 project to 1.3.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-131-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.1 project to 1.3.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-131-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.2', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.1 project to 1.3.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-131-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.1 project to 1.3.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-131-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.1 project to 1.3.2', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-131-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.2', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 1);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.2 project to 1.3.4', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-132-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.4', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 12);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.2 project to 1.3.4', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-132-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.4', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 12);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.2 project to 1.3.4', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-132-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.4', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 13);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.2 project to 1.3.4', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-132-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.4', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 13);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.2 project to 1.3.4', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-132-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.3.4', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 12);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.3.4 project to 1.4.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-134-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 19);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.3.4 project to 1.4.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-134-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 26);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.3.4 project to 1.4.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-134-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 19);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.3.4 project to 1.4.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-134-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 19);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.3.4 project to 1.4.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-134-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 25);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.4.0 project to 1.4.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-140-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.4.0 project to 1.4.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-140-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.4.0 project to 1.4.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-140-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.4.0 project to 1.4.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-140-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.4.0 project to 1.4.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-140-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 8);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.4.1 project to 1.5.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-141-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 26);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.4.1 project to 1.5.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-141-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 26);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.4.1 project to 1.5.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-141-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 26);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.4.1 project to 1.5.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-141-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 26);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.4.1 project to 1.5.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-141-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 25);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.0 project to 1.5.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-150-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.0 project to 1.5.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-150-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.1', debug: true, output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[3];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.5.0 project to 1.5.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-150-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.5.0 project to 1.5.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-150-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 9);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.5.0 project to 1.5.1', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-150-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.5.1', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 8);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-nolib'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project using MSGraphClient to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-nolib-graph'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 18);
    });
  });

  it('e2e: shows correct number of findings for upgrading no framework web part 1.5.1 project using AadHttpClient to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-nolib-aad'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 17);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.1 project to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading react web part 1.5.1 project using MSGraphClient to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-react-graph'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 21);
    });
  });

  it('e2e: shows correct number of findings for upgrading application customizer 1.5.1 project to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-applicationcustomizer'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading list view command set 1.5.1 project to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-listviewcommandset'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 16);
    });
  });

  it('e2e: shows correct number of findings for upgrading field customizer react 1.5.1 project to 1.6.0', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0', output: 'json' } }, (err?: any) => {
      const findings: Finding[] = log[0];
      assert.equal(findings.length, 15);
    });
  });

  it('shows all information with output format json', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-fieldcustomizer-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      assert(JSON.stringify(log[0]).indexOf('"resolution":') > -1);
    });
  });

  it('returns markdown report with output format md', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-react-graph'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'md', toVersion: '1.6.0' } }, (err?: any) => {
      assert(log[0].indexOf('## Findings') > -1);
    });
  });

  it('returns text report with output format default', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-151-webpart-react-graph'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.6.0' } }, (err?: any) => {
      assert(log[0].indexOf('Execute in command line') > -1);
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.PROJECT_UPGRADE));
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