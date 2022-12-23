import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./web-applicationcustomizer-add');
import * as SpoCustomActionAddCommand from '../customaction/customaction-add';


describe(commands.WEB_APPLICATIONCUSTOMIZER_ADD, () => {
  const webUrl = 'https://contoso.sharepoint.com';
  const title = 'PageFooter';
  const clientSideComponentId = '76d5f8c8-6228-4df8-a2da-b94cbc8115bc';
  const clientSideComponentProperties = '';


  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
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
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
  });

  afterEach(() => {
    sinonUtil.restore([
      Cli.executeCommand
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.WEB_APPLICATIONCUSTOMIZER_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds application customizer to a specific site without specifying clientSideComponentId', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake(async (command) => {
      if (command === SpoCustomActionAddCommand) {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId } } as any);
    assert(loggerLogToStderrSpy.notCalled);
  });

  it('adds application customizer to a specific site while specifying clientSideComponentId', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake(async (command, args) => {
      if (command === SpoCustomActionAddCommand && args.options["clientSideComponentProperties"] === '') {
        return;
      }

      throw 'Invalid request';
    });

    await command.action(logger, { options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, verbose: true } } as any);
    assert(loggerLogToStderrSpy.called);
  });

  it('throws an error when error occurs on adding the application customizer', async () => {
    sinon.stub(Cli, 'executeCommand').callsFake(async (command) => {
      if (command === SpoCustomActionAddCommand) {
        throw 'Error occured.';
      }

      throw 'Invalid request';
    });

    await assert.rejects(command.action(logger, { options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, verbose: true } } as any), new CommandError('Error occured.'));
  });

  it('fails validation if the webUrl option is not a valid SharePoint site URL', async () => {
    const actual = await command.validate({ options: { webUrl: 'foo', title: title, clientSideComponentId: clientSideComponentId } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the clientSideComponentId option is not a valid GUID', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if all options are passed', async () => {
    const actual = await command.validate({ options: { webUrl: webUrl, title: title, clientSideComponentId: clientSideComponentId, clientSideComponentProperties: clientSideComponentProperties } }, commandInfo);
    assert.strictEqual(actual, true);
  });
}); 
