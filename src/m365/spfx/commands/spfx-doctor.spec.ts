import * as assert from 'assert';
import * as child_process from 'child_process';
import * as sinon from 'sinon';
import { SinonSandbox } from 'sinon';
import appInsights from '../../../appInsights';
import { Cli, CommandInfo, Logger } from '../../../cli';
import Command, { CommandError } from '../../../Command';
import { sinonUtil } from '../../../utils';
import commands from '../commands';
const command: Command = require('./spfx-doctor');

describe(commands.DOCTOR, () => {
  let log: string[];
  let sandbox: SinonSandbox;
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  const packageVersionResponse = (name: string, version: string): string => {
    return `{
      "dependencies": {
        "${name}": {
          "version": "${version}"
        }
      }
    }`;
  };
  const getStatus = (status: number, message: string): string => {
    return (<any>command).getStatus(status, message);
  };

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    loggerLogSpy = sinon.spy(logger, 'log');
    sinon.stub(process, 'platform').value('linux');
  });

  afterEach(() => {
    sinonUtil.restore([
      sandbox,
      child_process.exec,
      process.platform
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.DOCTOR), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('passes all checks for SPFx v1.11 project when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.22.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.22.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.11 project when all requirements met (debug)', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.11 generator installed locally when all requirements met (debug)', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.11 generator installed globally when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? packageVersionResponse(packageName, '1.11.0') : '{ }');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.11 generator installed locally when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.10 project when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.10 project when all requirements met (debug)', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.10 generator installed locally when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.10 generator installed locally when all requirements met (debug)', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? '{ }' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes all checks for SPFx v1.10 generator installed globally when all requirements met', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const args = file.split(' ');
      const packageName: string = args[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{ }');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, args[args.length - 1] === '-g' ? packageVersionResponse(packageName, '1.10.0') : '{ }');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')), 'Invalid gulp-cli version reported');
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with error when SPFx not found', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        case '@microsoft/generator-sharepoint':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } } as any, (err: any) => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('SharePoint Framework not found')));
        assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with error when SPFx not found (debug)', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        case '@microsoft/generator-sharepoint':
          callback(undefined, '{}');
          return {} as child_process.ChildProcess;
        default:
          callback(new Error(`${file} ENOENT`));
          return {} as child_process.ChildProcess;
      }
    });

    command.action(logger, {options: {debug: true } } as any, (err: any) => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('SharePoint Framework not found')));
        assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SPO compatibility check for SPFx v1.11.0', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'spo' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SPO compatibility check for SPFx v1.10.0', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'spo' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SP2019 compatibility check for SPFx v1.4.1', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.4.1'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'sp2019' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SP2019')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if output does not equal text.', async () => {
    const actual = await command.validate({
      options: {
        output: 'json'
      }
    }, commandInfo);

    assert.notStrictEqual(actual, true);
  });

  it('fails SP2019 compatibility check for SPFx v1.5.0', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.5.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'sp2019' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'Not supported in SP2019')));
        assert(loggerLogSpy.calledWith('- Use SharePoint Framework v1.4.1'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SP2016 compatibility check for SPFx v1.1.0', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.1.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'sp2016' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Supported in SP2016')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails SP2016 compatibility check for SPFx v1.2.0', (done) => {
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.2.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false, env: 'sp2016' } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'Not supported in SP2016')));
        assert(loggerLogSpy.calledWith('- Use SharePoint Framework v1.1'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes Node check when version meets single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return { } as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v10.18.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes Node check when version meets double range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return { } as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'Node v8.0.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails Node check when version does not meet single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v12.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return { } as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^10 required')));
        assert(loggerLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails Node check when version does not meet double range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v12.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return { } as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^8 || ^10 required')));
        assert(loggerLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with friendly error message when npm not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.0.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      callback(new Error(`${file} ENOENT`));
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } } as any, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('npm not found')));
        assert(!loggerLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
  
  it('passes yo check when yo found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'yo v3.1.1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails yo check when yo not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'yo not found')));
        assert(loggerLogSpy.calledWith('- npm i -g yo@3'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes gulp-cli check when gulp-cli found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');

    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        case 'gulp-cli':
          callback(undefined, packageVersionResponse(packageName, '2.3.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`), '', '');
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'gulp-cli v2.3.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails gulp-cli check when gulp-cli not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');

    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        default:
          callback(new Error(`${file} ENOENT`), '', '');
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'gulp-cli not found')));
        assert(loggerLogSpy.calledWith('- npm i -g gulp-cli@2'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails gulp check when gulp is found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'gulp':
          callback(null, packageVersionResponse(packageName, '4.0.0'), '');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'gulp should be removed')));
        assert(loggerLogSpy.calledWith('- npm un -g gulp'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes typescript check when typescript not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'typescript':
          callback(undefined, '{ }');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'bundled typescript used')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails typescript check when typescript found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'typescript':
          callback(undefined, packageVersionResponse(packageName, '3.7.5'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(1, 'typescript v3.7.5 installed in the project')));
        assert(loggerLogSpy.calledWith('- npm un typescript'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns error when used with an unsupported version of spfx', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '0.9.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } } as any, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`spfx doctor doesn't support SPFx v0.9.0 at this moment`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('uses alternative symbols for win32', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sandbox.stub(process, 'platform').value('win32');
    if (process.env.CI) {
      sandbox.stub(process.env, 'CI').value('false');
    }
    if (process.env.TERM) {
      sandbox.stub(process.env, 'TERM').value('16');
    }
    sinon.stub(child_process, 'exec').callsFake((file, callback: any) => {
      const packageName: string = file.split(' ')[2];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        default:
          callback({message: `${file} ENOENT` } as any, '', '');
      }
      return { } as child_process.ChildProcess;
    });

    command.action(logger, {options: {debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying environment', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '-e, --env [env]') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types, 'undefined', 'command types undefined');
    assert.notStrictEqual(command.types.string, 'undefined', 'command string types undefined');
  });

  it('configures env as string option', () => {
    const types = command.types;
    ['e', 'env'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('passes validation when no options specified', async () => {
    const actual = await command.validate({options: { } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2016 env specified', async () => {
    const actual = await command.validate({options: {env: 'sp2016' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2019 env specified', async () => {
    const actual = await command.validate({options: {env: 'sp2019' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation when spo env specified', async () => {
    const actual = await command.validate({options: {env: 'spo' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when 2016 env specified', async () => {
    const actual = await command.validate({options: {env: '2016' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });
});