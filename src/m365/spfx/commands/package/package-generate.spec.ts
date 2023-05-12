import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { fsUtil } from '../../../../utils/fsUtil';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
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
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    (command as any).archive = admZipMock;
    commandInfo = Cli.getCommandInfo(command);
    Cli.getInstance().config;
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
    sinon.stub(fsUtil, 'readdirR').callsFake(_ => ['file1.png', 'file.json']);
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
    sinonUtil.restore([
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
      fsUtil.copyRecursiveSync,
      fsUtil.readdirR
    ]);
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.PACKAGE_GENERATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('creates a package for the specified HTML snippet', async () => {
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert(archiveWriteZipSpy.called);
  });

  it('creates a package for the specified HTML snippet (debug)', async () => {
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all',
        debug: true
      }
    });
    assert(archiveWriteZipSpy.called);
  });

  it('creates a package exposed as a Teams tab', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab']).replace(/"/g, '&quot;')));
  });

  it('creates a package exposed as a Teams personal app', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'personalApp'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
  });

  it('creates a package exposed as a Teams tab and personal app', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$supportedHosts$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', JSON.stringify(['SharePointWebPart', 'TeamsTab', 'TeamsPersonalApp']).replace(/"/g, '&quot;')));
  });

  it('handles exception when creating a temp folder failed', async () => {
    sinonUtil.restore(fs.mkdtempSync);
    sinon.stub(fs, 'mkdtempSync').throws(new Error('An error has occurred'));
    const archiveWriteZipSpy = sinon.spy(admZipMock, 'writeZip');
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }), (err) => err === 'An error has occurred');
    assert(archiveWriteZipSpy.notCalled);
  });

  it('handles error when creating the package failed', async () => {
    sinon.stub(admZipMock, 'writeZip').throws(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }), (err) => err === 'An error has occurred');
  });

  it('removes the temp directory after the package has been created', async () => {
    sinonUtil.restore(fs.rmdirSync);
    const fsrmdirSyncSpy = sinon.stub(fs, 'rmdirSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    });
    assert(fsrmdirSyncSpy.called);
  });

  it('removes the temp directory if creating the package failed', async () => {
    sinonUtil.restore(fs.rmdirSync);
    const fsrmdirSyncSpy = sinon.stub(fs, 'rmdirSync').callsFake(_ => { });
    sinon.stub(admZipMock, 'writeZip').throws(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }));
    assert(fsrmdirSyncSpy.called);
  });

  it('prompts user to remove the temp directory manually if removing it automatically failed', async () => {
    sinonUtil.restore(fs.rmdirSync);
    sinon.stub(fs, 'rmdirSync').throws(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }), (err) => err === 'An error has occurred while removing the temp folder at /tmp/abc. Please remove it manually.');
  });

  it('leaves unknown token as-is', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$token$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '$token$'));
  });

  it('exposes page context globally', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$exposePageContextGlobally$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposePageContextGlobally: true
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
  });

  it('exposes Teams context globally', async () => {
    sinonUtil.restore([fs.readFileSync, fs.writeFileSync]);
    sinon.stub(fs, 'readFileSync').callsFake(_ => '$exposeTeamsContextGlobally$');
    const fsWriteFileSyncSpy = sinon.stub(fs, 'writeFileSync').callsFake(_ => { });
    await command.action(logger, {
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam',
        name: 'amsterdam-weather',
        html: 'abc',
        allowTenantWideDeployment: true,
        exposeTeamsContextGlobally: true
      }
    });
    assert(fsWriteFileSyncSpy.calledWith('file.json', '!0'));
  });

  it(`fails validation if the enableForTeams option is invalid`, async () => {
    const actual = await command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to tab`, async () => {
    const actual = await command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'tab'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to personalApp`, async () => {
    const actual = await command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'personalApp'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it(`passes validation if the enableForTeams option is set to all`, async () => {
    const actual = await command.validate({
      options: {
        webPartTitle: 'Amsterdam weather',
        webPartDescription: 'Shows weather in Amsterdam', name: 'amsterdam-weather',
        html: '@amsterdam-weather.html', allowTenantWideDeployment: true,
        enableForTeams: 'all'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });
});
