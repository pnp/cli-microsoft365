import request from '../../../../request';
import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./project-externalize');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import { Project, ExternalConfiguration } from './project-upgrade/model';

describe(commands.PROJECT_EXTERNALIZE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  const logEntryToCheck = 1; //necessary as long as we display the beta message

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
      log: (msg: string) => {
        log.push(msg);
      }
    };
    telemetry = null;
    (command as any).allFindings = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      (command as any).getProjectRoot,
      (command as any).getProjectVersion,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync,
      request.head
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.PROJECT_EXTERNALIZE), true);
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
      assert.equal(telemetry.name, commands.PROJECT_EXTERNALIZE);
    });
  });

  it('shows error if the project path couldn\'t be determined', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
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
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
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
      assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, 3)));
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

  //#region findings

  it('e2e: shows correct number of findings for externalizing react web part 1.8.2 project', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('package.json') && path.indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      } else if (path.endsWith('package.json') && path.indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: true } }, (err?: any) => {
      const findings: ExternalConfiguration = JSON.parse(log[logEntryToCheck + 3]); //because debug is enabled
      assert.equal(findings['@pnp/pnpjs'].path, 'https://unpkg.com/@pnp/pnpjs@1.3.5/dist/pnpjs.es5.umd.bundle.min.js');
    });
  });

  it('handles failures properly', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('package.json') && path.indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      } else if (path.endsWith('package.json') && path.indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    const originalWriteFileSync = fs.writeFileSync;
    sinon.stub(fs, 'writeFileSync').callsFake((path: string, value: string, encoding: string) => {
      if(path.endsWith('report.json')) {
        throw new Error('file is locked');
      } else {
        return originalWriteFileSync(path, value, encoding);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: true, outputFile: 'report.json' } }, (err?: any) => {
      assert.equal(!err, false);
    });
  });
  //#endregion

  it('shows all information with output format json', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      assert(log[logEntryToCheck].startsWith('{'));
    });
  });

  it('returns markdown report with output format md', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'md', } }, (err?: any) => {
      assert(log[logEntryToCheck].indexOf('## Findings') > -1);
    });
  });

  it('returns text report with output format default', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path: string) => {
      if (path.endsWith('package.json') && path.indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      } else if (path.endsWith('package.json') && path.indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { } }, (err?: any) => {
      assert(JSON.stringify(log[logEntryToCheck]).startsWith('['));
    });
  });

  it('writes upgrade report to file when outputFile specified', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));
    const writeFileSyncSpy: sinon.SinonStub = sinon.stub(fs, 'writeFileSync').callsFake(() => { });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'md', outputFile: '/foo/report.md' } }, (err?: any) => {
      assert(writeFileSyncSpy.called);
    });
  });

  it('writes JSON upgrade report to file when outputFile specified in json output mode', () => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/o365/spfx/commands/project/project-upgrade/test-projects/spfx-182-webpart-react'));
    let typeofReport: string = '';
    sinon.stub(fs, 'writeFileSync').callsFake((path: string, contents: any) => {
      typeofReport = typeof contents;
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', outputFile: '/foo/report.md' } }, (err?: any) => {
      assert.equal(typeofReport, 'string');
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

  it('passes validation when package manager not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, true);
  });
  
  it('fails validation when non-existent path specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = (command.validate() as CommandValidate)({ options: { outputFile: '/foo/file.md' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when valid file path specified', () => {
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    const actual = (command.validate() as CommandValidate)({ options: { outputFile: '/foo/file.md' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.PROJECT_EXTERNALIZE));
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