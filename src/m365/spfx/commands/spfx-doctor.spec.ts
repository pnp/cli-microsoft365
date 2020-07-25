import commands from '../commands';
import Command, { CommandError, CommandOption, CommandValidate, CommandTypes } from '../../../Command';
import * as sinon from 'sinon';
import { SinonSandbox } from 'sinon';
import appInsights from '../../../appInsights';
const command: Command = require('./spfx-doctor');
import * as assert from 'assert';
import Utils from '../../../Utils';
import * as child_process from 'child_process';

describe(commands.DOCTOR, () => {
  let log: string[];
  let sandbox: SinonSandbox;
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
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
  }

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    sinon.stub(process, 'platform').value('linux');
  });

  afterEach(() => {
    Utils.restore([
      sandbox,
      child_process.execFile,
      process.platform
    ]);
  });

  after(() => {
    Utils.restore([
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.14.6');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '4.0.2'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.22.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.14.6')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v4.0.2')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.14.6');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '4.0.2'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.14.6')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v4.0.2')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? '{}' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? packageVersionResponse(packageName, '1.11.0') : '{}');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.14.6');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? '{}' : packageVersionResponse(packageName, '1.11.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.11.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.14.6')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? '{}' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? '{}' : packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, '{}');
          break;
        case '@microsoft/generator-sharepoint':
          callback(undefined, (args as string[])[(args as string[]).length - 1] === '-g' ? packageVersionResponse(packageName, '1.10.0') : '{}');
          break;
        case 'yo':
          callback(undefined, packageVersionResponse(packageName, '3.1.1'));
          break;
        case 'gulp':
          callback(undefined, packageVersionResponse(packageName, '3.9.1'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')), 'Invalid Node version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.13.4')), 'Invalid npm version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')), 'Invalid yo version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')), 'Invalid gulp version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'Invalid react version reported');
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')), 'Invalid typescript reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with error when SPFx not found', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
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

    cmdInstance.action({ options: { debug: false } }, (err: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('SharePoint Framework not found')));
        assert(!cmdInstanceLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails with error when SPFx not found (debug)', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
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

    cmdInstance.action({ options: { debug: true } }, (err: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'SharePoint Framework')), 'SharePoint Framework found');
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('SharePoint Framework not found')));
        assert(!cmdInstanceLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SPO compatibility check for SPFx v1.11.0', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.11.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'spo' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SPO compatibility check for SPFx v1.10.0', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'spo' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Supported in SPO')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SP2019 compatibility check for SPFx v1.4.1', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.4.1'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'sp2019' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Supported in SP2019')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails SP2019 compatibility check for SPFx v1.5.0', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.5.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'sp2019' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'Not supported in SP2019')));
        assert(cmdInstanceLogSpy.calledWith('- Use SharePoint Framework v1.4.1'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes SP2016 compatibility check for SPFx v1.1.0', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.1.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'sp2016' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Supported in SP2016')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails SP2016 compatibility check for SPFx v1.2.0', (done) => {
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.2.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false, env: 'sp2016' } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'Not supported in SP2016')));
        assert(cmdInstanceLogSpy.calledWith('- Use SharePoint Framework v1.1'), 'No fix provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v10.18.0')));
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'Node v8.0.0')));
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^10.0.0 required')));
        assert(cmdInstanceLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.9.0'));
          return {} as child_process.ChildProcess;
      }

      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'Node v12.0.0 found, v^8.0.0 || ^10.0.0 required')));
        assert(cmdInstanceLogSpy.calledWith('- Install Node.js v10'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes npm check when version meets single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '5.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.6.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v5.0.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes npm.cmd check when os is Windows', (done) => {
    Utils.restore(process.platform);
    sinon.stub(process, 'platform').value('win32');

    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm.cmd' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '5.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.6.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v5.0.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes npm check when version meets double range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'npm v6.0.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails npm check when version does not meet single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '4.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.6.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'npm v4.0.0 found, v^5.0.0 required')));
        assert(cmdInstanceLogSpy.calledWith('- npm i -g npm@5'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails npm check when version does not meet double range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '7.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'npm v7.0.0 found, v^5.0.0 || ^6.0.0 required')));
        assert(cmdInstanceLogSpy.calledWith('- npm i -g npm@6'), 'No fix provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      callback(new Error(`${file} ENOENT`));
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, (err: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('npm not found')));
        assert(!cmdInstanceLogSpy.calledWith('Recommended fixes:'), 'Fixes provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
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
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'yo v3.1.1')));
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'yo not found')));
        assert(cmdInstanceLogSpy.calledWith('- npm i -g yo'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes gulp check when gulp found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');

    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(null, '6.13.4', '');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        case 'gulp':
          callback(null, packageVersionResponse(packageName, '3.9.1'), '');
          break;
        default:
          callback(new Error(`${file} ENOENT`), '', '');
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'gulp v3.9.1')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails gulp check when gulp not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'gulp not found')));
        assert(cmdInstanceLogSpy.calledWith('- npm i -g gulp'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes react check when react not found', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(!cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')), 'react found');
        assert(!cmdInstanceLogSpy.calledWith(getStatus(1, 'react not found, v16.8.5 required')), 'react not found');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes react check when react meets single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '5.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.6.0'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '15.0.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v15.0.0')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes react check when react meets specific version prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.5'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'react v16.8.5')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails react check when react does not meet single range prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v8.0.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '5.0.0');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.6.0'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.0.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'react v16.0.0 found, v^15 required')));
        assert(cmdInstanceLogSpy.calledWith('- npm i react@15'), 'No fix provided');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails react check when react does not meet specific version prerequisite', (done) => {
    const sandbox = sinon.createSandbox();
    sandbox.stub(process, 'version').value('v10.18.0');
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'react':
          callback(undefined, packageVersionResponse(packageName, '16.8.6'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'react v16.8.6 found, v16.8.5 required')));
        assert(cmdInstanceLogSpy.calledWith('- npm i react@16.8.5'), 'No fix provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '1.10.0'));
          break;
        case 'typescript':
          callback(undefined, '{}');
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'bundled typescript used')));
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
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
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(1, 'typescript v3.7.5 installed in the project')));
        assert(cmdInstanceLogSpy.calledWith('- npm un typescript'), 'No fix provided');
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(undefined, '6.13.4');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(undefined, packageVersionResponse(packageName, '0.9.0'));
          break;
        default:
          callback(new Error(`${file} ENOENT`));
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, (err: any) => {
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
    sinon.stub(child_process, 'execFile').callsFake((file, args, callback: any) => {
      if (file === 'npm' && args && args.length === 1 && args[0] === '-v') {
        callback(null, '6.13.4', '');
        return {} as child_process.ChildProcess;
      }

      const packageName: string = (args as string[])[1];
      switch (packageName) {
        case '@microsoft/sp-core-library':
          callback(null, packageVersionResponse(packageName, '1.10.0'), '');
          break;
        default:
          callback({ message: `${file} ENOENT` } as any, '', '');
      }
      return {} as child_process.ChildProcess;
    });

    cmdInstance.action({ options: { debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(getStatus(0, 'SharePoint Framework v1.10.0')), 'Invalid SharePoint Framework version reported');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports specifying environment', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '-e, --env [env]') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('configures command types', () => {
    assert.notStrictEqual(typeof command.types(), 'undefined', 'command types undefined');
    assert.notStrictEqual((command.types() as CommandTypes).string, 'undefined', 'command string types undefined');
  });

  it('configures env as string option', () => {
    const types = (command.types() as CommandTypes);
    ['e', 'env'].forEach(o => {
      assert.notStrictEqual((types.string as string[]).indexOf(o), -1, `option ${o} not specified as string`);
    });
  });

  it('passes validation when no options specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2016 env specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { env: 'sp2016' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when sp2019 env specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { env: 'sp2019' } });
    assert.strictEqual(actual, true);
  });

  it('passes validation when spo env specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { env: 'spo' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation when 2016 env specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { env: '2016' } });
    assert.notStrictEqual(actual, true);
  });
});