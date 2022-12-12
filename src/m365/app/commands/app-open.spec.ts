import * as assert from 'assert';
import * as sinon from 'sinon';
import * as fs from 'fs';
import { telemetry } from '../../../telemetry';
import auth from '../../../Auth';
import { Cli } from '../../../cli/Cli';
import { CommandInfo } from '../../../cli/CommandInfo';
import { Logger } from '../../../cli/Logger';
import Command, { CommandError } from '../../../Command';
import { pid } from '../../../utils/pid';
import { sinonUtil } from '../../../utils/sinonUtil';
import * as open from 'open';
import commands from '../commands';
const command: Command = require('./app-open');

describe(commands.OPEN, () => {
  let log: string[];
  let logger: Logger;
  let cli: Cli;
  let openStub: sinon.SinonStub;
  let getSettingWithDefaultValueStub: sinon.SinonStub;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'readFileSync').callsFake(() => JSON.stringify({
      "apps": [
        {
          "appId": "9b1b1e42-794b-4c71-93ac-5ed92488b67f",
          "name": "CLI app1"
        }
      ]
    }));
    commandInfo = Cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    cli = Cli.getInstance();
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
    (command as any)._open = open;
    openStub = sinon.stub(command as any, '_open').callsFake(() => Promise.resolve(null));
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => false));
  });

  afterEach(() => {
    openStub.restore();
    getSettingWithDefaultValueStub.restore();
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName,
      fs.existsSync,
      fs.readFileSync
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.OPEN), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the appId is not a valid guid', async () => {
    const actual = await command.validate({ options: { appId: 'abc' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if valid appId-guid is specified', async () => {
    const actual = await command.validate({ options: { appId: '9b1b1e42-794b-4c71-93ac-5ed92488b67f' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('shows message with url when the app specified with the appId is found', async () => {
    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await command.action(logger, {
      options: {
        appId: appId
      }
    });
    assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });

  it('shows message with url when the app specified with the appId is found (verbose)', async () => {
    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await command.action(logger, {
      options: {
        verbose: true,
        appId: appId
      }
    });
    assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });

  it('shows message with preview-url when the app specified with the appId is found', async () => {
    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await command.action(logger, {
      options: {
        appId: appId,
        preview: true
      }
    });
    assert(loggerLogSpy.calledWith(`Use a web browser to open the page https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });

  it('shows message with url when the app specified with the appId is found (using autoOpenInBrowser)', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => true));

    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await command.action(logger, {
      options: {
        appId: appId
      }
    });
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });

  it('shows message with preview-url when the app specified with the appId is found (using autoOpenInBrowser)', async () => {
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => true));

    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await command.action(logger, {
      options: {
        appId: appId,
        preview: true
      }
    });
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });

  it('throws error when open in browser fails', async () => {
    openStub.restore();
    openStub = sinon.stub(command as any, '_open').callsFake(() => Promise.reject("An error occurred"));
    getSettingWithDefaultValueStub.restore();
    getSettingWithDefaultValueStub = sinon.stub(cli, 'getSettingWithDefaultValue').callsFake((() => true));

    const appId = "9b1b1e42-794b-4c71-93ac-5ed92488b67f";
    await assert.rejects(command.action(logger, {
      options: {
        appId: appId,
        preview: true
      }
    }), new CommandError("An error occurred"));
    assert(loggerLogSpy.calledWith(`Opening the following page in your browser: https://preview.portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/Overview/appId/${appId}/isMSAApp/`));
  });
});
