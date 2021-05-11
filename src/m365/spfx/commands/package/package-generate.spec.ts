import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./package-generate');

const admZipMock = {
  // we need these unused params so that they can be properly mocked with sinon
  /* eslint-disable @typescript-eslint/no-unused-vars */
  addFile: (entryName: string, data: Buffer, comment?: string, attr?: number) => { },
  addLocalFile: (localPath: string, zipPath?: string, zipName?: string) => { },
  writeZip: (targetFileName?: string, callback?: (error: Error | null) => void) => { }
  /* eslint-enable @typescript-eslint/no-unused-vars */
};

describe(commands.PACKAGE_GENERATE, () => {
  let log: any[];
  let logger: Logger;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    (command as any).archive = admZipMock;
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
    sinon.stub(fs, 'mkdtempSync').callsFake(_ => '/tmp/abc');
    sinon.stub(Utils, 'readdirR').callsFake(_ => ['file1.png', 'file.json']);
    sinon.stub(fs, 'readFileSync').callsFake(_ => 'abc');
    sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    sinon.stub(fs, 'rmdirSync').callsFake(_ => { });
    sinon.stub(fs, 'mkdirSync').callsFake(_ => '/tmp/abc/def');
    sinon.stub(fs, 'copyFileSync').callsFake(_ => { });
    sinon.stub(fs, 'statSync').callsFake(src => {
      return {
        isDirectory: () => src.toString().indexOf('.') < 0
      } as any;
    });
  });

  afterEach(() => {
    Utils.restore([
      (command as any).generateNewId,
      admZipMock.addFile,
      admZipMock.addLocalFile,
      admZipMock.writeZip,
      fs.copyFileSync,
      fs.mkdtempSync,
      fs.mkdirSync,
      fs.readFileSync,
      fs.rmdirSync,
      fs.statSync,
      fs.writeFileSync,
      Utils.copyRecursiveSync,
      Utils.readdirR
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PACKAGE_GENERATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a package for the specified HTML snippet', done => {
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(archiveWriteZipSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a package for the specified HTML snippet (debug)', done => {
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(archiveWriteZipSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a package exposed as a Teams tab', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab',
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab']).replace(/"/g, '&quot;')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a package exposed as a Teams personal app', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'personalApp',
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('creates a package exposed as a Teams tab and personal app', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles exception when creating a temp folder failed', done => {
    Utils.restore(fs.mkdtempSync);
    sinon.stub(fs, 'mkdtempSync').throws(new Error('An error has occurred'));
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, 'An error has occurred');
        assert(archiveWriteZipSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error when creating the package failed', done => {
    sinon.stub(admZipMock, 'writeZip').throws(new Error('An error has occurred'));
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, 'An error has occurred');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the temp directory after the package has been created', done => {
    Utils.restore(fs.rmdirSync);
    const fsrmdirSyncSpy = sinon.stub(fs, 'rmdirSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        assert(fsrmdirSyncSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('removes the temp directory if creating the package failed', done => {
    Utils.restore(fs.rmdirSync);
    const fsrmdirSyncSpy = sinon.stub(fs, 'rmdirSync').callsFake(_ => { });
    sinon.stub(admZipMock, 'writeZip').throws(new Error('An error has occurred'));
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.notStrictEqual(typeof err, 'undefined');
        assert(fsrmdirSyncSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts user to remove the temp directory manually if removing it automatically failed', done => {
    Utils.restore(fs.rmdirSync);
    sinon.stub(fs, 'rmdirSync').throws(new Error('An error has occurred'));
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: false
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(err, 'An error has occurred while removing the temp folder at /tmp/abc. Please remove it manually.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('leaves unknown token as-is', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$token$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab',
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', '$token$'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exposes page context globally', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$exposePageContextGlobally$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposePageContextGlobally: true,
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('exposes Teams context globally', done => {
    Utils.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$exposeTeamsContextGlobally$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        packageName: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposeTeamsContextGlobally: true,
        debug: false
      }
    }, () => {
      try {
        assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it(`fails validation if the enableForTeams option is invalid`, () => {
    const actual = command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', packageName: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to tab`, () => {
    const actual = command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', packageName: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    });
    assert.strictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to personalApp`, () => {
    const actual = command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', packageName: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'personalApp'
      }
    });
    assert.strictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to all`, () => {
    const actual = command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', packageName: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});