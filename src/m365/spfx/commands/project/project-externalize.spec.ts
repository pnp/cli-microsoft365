import request from '../../../../request';
import commands from '../../commands';
import Command, { CommandOption, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
const command: Command = require('./project-externalize');
import * as assert from 'assert';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../Utils';
import { Project, ExternalConfiguration, External } from './model';
import { ExternalizeEntry, FileEdit } from './project-externalize/';
import * as requestNative from 'request-promise-native';

describe(commands.PROJECT_EXTERNALIZE, () => {
  let vorpal: Vorpal;
  let log: any[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;
  const logEntryToCheck = 1; //necessary as long as we display the beta message
  const projectPath: string = './src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react';

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
      request.head,
      request.post,
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

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.PROJECT_EXTERNALIZE);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('shows error if the project path couldn\'t be determined', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => null);

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('searches for package.json in the parent folder when it doesn\'t exist in the current folder', (done) => {
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('package.json')) {
        return false;
      }
      else {
        return true;
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Couldn't find project root folder`, 1)));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('correctly handles the case when .yo-rc.json exists but doesn\'t contain spfx project info', (done) => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json') || path.toString().endsWith('package.json')) {
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
      else if (path.toString().endsWith('package.json')) {
        return JSON.stringify({
          dependencies: {
            '@microsoft/sp-core-library': '1.8.1'
          }
        });
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(getProjectVersionSpy.lastCall.returnValue, '1.8.1');
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('determines the current version from .yo-rc.json when available', (done) => {
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
        return `{
          "@microsoft/generator-sharepoint": {
            "version": "0.4.1",
            "libraryName": "spfx-041",
            "libraryId": "dd1a0a8d-e043-4ca0-b9a4-256e82a66177",
            "environment": "spo"
          }
        }`;
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    const getProjectVersionSpy = sinon.spy(command as any, 'getProjectVersion');

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(getProjectVersionSpy.lastCall.returnValue, '0.4.1');
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('tries to determine the current version from package.json if .yo-rc.json doesn\'t exist', (done) => {
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
          "name": "spfx-041",
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('shows error if the project version couldn\'t be determined', (done) => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('.yo-rc.json')) {
        return false;
      }
      else {
        return originalExistsSync(path);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Unable to determine the version of the current SharePoint Framework project`, 3)));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('determining project version doesn\'t fail if .yo-rc.json is empty', () => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1' } }, (err?: any) => {
      assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, '1.4.1');
    });
  });

  it('determining project version doesn\'t fail if package.json is empty', () => {
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

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { toVersion: '1.4.1' } }, (err?: any) => {
      assert.strictEqual(getProjectVersionSpy.lastCall.returnValue, undefined);
    });
  });

  it('loads config.json when available', () => {
    const originalExistsSync = fs.existsSync;
    sinon.stub(fs, 'existsSync').callsFake((path) => {
      if (path.toString().endsWith('config.json')) {
        return true;
      }
      else {
        return originalExistsSync(path);
      }
    });
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('config.json')) {
        return '{}';
      }
      else {
        return originalReadFileSync(path, options);
      }
    });

    const getProject = (command as any).getProject;
    const project: Project = getProject(projectPath);
    assert.notEqual(typeof (project.configJson), 'undefined');
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
    assert.equal(typeof (project.packageJson), 'undefined');
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
    assert.equal(typeof (project.tsConfigJson), 'undefined');
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
    assert.equal(typeof (project.configJson), 'undefined');
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
    assert.equal(typeof (project.packageJson), 'undefined');
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
    assert.equal(typeof (project.yoRcJson), 'undefined');
  });

  //#region findings

  it('e2e: shows correct number of findings for externalizing react web part 1.8.2 project', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && path.toString().indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve(JSON.stringify({ scriptType: 'module' })));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: true } }, (err?: any) => {
      try {
        const findings: { externalConfiguration: { externals: ExternalConfiguration }, edits: FileEdit[] } = log[logEntryToCheck + 3]; //because debug is enabled
        assert.equal((findings.externalConfiguration.externals['@pnp/pnpjs'] as unknown as External).path, 'https://unpkg.com/@pnp/pnpjs@1.3.5/dist/pnpjs.es5.umd.min.js');
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('returns edit suggestions', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && path.toString().indexOf('logging') > -1) {
        return JSON.stringify({
          main: "./dist/logging.es5.umd.bundle.js",
          module: "./dist/logging.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('common') > -1) {
        return JSON.stringify({
          main: "./dist/common.es5.umd.bundle.js",
          module: "./dist/common.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/logging'] = '1.3.5';
        pConfig.dependencies['@pnp/common'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.resolve(JSON.stringify({ scriptType: 'script' }));
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: false } }, (err?: any) => {
      try {
        const findings: { externalConfiguration: { externals: ExternalConfiguration }, edits: FileEdit[] } = log[0];
        assert.notEqual(findings.edits.length, 0);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('handles failures properly', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && path.toString().indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('tntjs') > -1) {
        return JSON.stringify({
          main: "./dist/tntjs.es5.umd.bundle.js",
          module: "./dist/tntjs.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('logging') > -1) {
        return JSON.stringify({
          main: "./dist/logging.es5.umd.bundle.js",
          module: "./dist/logging.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('common') > -1) {
        return JSON.stringify({
          main: "./dist/common.es5.umd.bundle.js",
          module: "./dist/common.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        pConfig.dependencies['@pnp/tntjs'] = '1.3.5';
        pConfig.dependencies['@pnp/logging'] = '1.3.5';
        pConfig.dependencies['@pnp/common'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake((options: requestNative.OptionsWithUrl) => {
      if ((options.body as string).indexOf('tnt') > -1) {
        return Promise.resolve(JSON.stringify({ scriptType: 'module' }));
      } else {
        return Promise.resolve(JSON.stringify({ scriptType: 'script' }));
      }
    });
    const originalWriteFileSync = fs.writeFileSync;
    sinon.stub(fs, 'writeFileSync').callsFake((path, value, encoding) => {
      if (path.toString().endsWith('report.json')) {
        throw new Error('file is locked');
      }
      else {
        return originalWriteFileSync(path, value, encoding);
      }
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json', debug: true } }, (err?: any) => {
      try {
        assert.notEqual(typeof err, 'undefined');
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });
  //#endregion

  it('outputs JSON object with output format json', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'json' } }, (err?: any) => {
      try {
        assert(JSON.stringify(log[0]).startsWith('{'));
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('returns markdown report with output format md', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { output: 'md', } }, (err?: any) => {
      try {
        assert(log[logEntryToCheck].indexOf('## Findings') > -1);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('returns text report with output format default', (done) => {
    sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-182-webpart-react'));
    const originalReadFileSync = fs.readFileSync;
    sinon.stub(fs, 'readFileSync').callsFake((path, options) => {
      if (path.toString().endsWith('package.json') && path.toString().indexOf('pnpjs') > -1) {
        return JSON.stringify({
          main: "./dist/pnpjs.es5.umd.bundle.js",
          module: "./dist/pnpjs.es5.umd.bundle.min.js"
        });
      }
      else if (path.toString().endsWith('package.json') && path.toString().indexOf('spfx-182-webpart-react') > -1) { //adding library on the fly so we get at least one result
        const pConfig = JSON.parse(originalReadFileSync(path, 'utf8'));
        pConfig.dependencies['@pnp/pnpjs'] = '1.3.5';
        return JSON.stringify(pConfig);
      }
      else {
        return originalReadFileSync(path, options);
      }
    });
    sinon.stub(request, 'head').callsFake(() => Promise.resolve());
    sinon.stub(request, 'post').callsFake(() => Promise.resolve(JSON.stringify({ scriptType: 'module' })));
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, (err?: any) => {
      try {
        assert.notEqual(log[1].indexOf('externalConfiguration'), -1);
        done();
      }
      catch (ex) {
        done(ex);
      }
    });
  });

  it('covers all text report branches', () => {
    const report = (command as any).serializeTextReport([
      {
        key: 'fake',
        path: 'https://fake.com/module.js',
        globalName: 'fakename',
        globalDependencies: ['fakeparent']
      } as ExternalizeEntry,
      {
        key: 'fakenoglobal',
        path: 'https://fake.com/module.js',
        globalDependencies: ['fakeparentnoglobal']
      } as ExternalizeEntry
    ]) as string;
    const emptyReport = (command as any).serializeTextReport([]) as string;
    assert(report.length > 87);

    // Windows processes JSON.stringify different then OSX/Linux and adds two empty characters
    assert(emptyReport.length === 122 || emptyReport.length === 124);
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
