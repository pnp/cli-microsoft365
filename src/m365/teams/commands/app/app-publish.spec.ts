import * as assert from 'assert';
import * as fs from 'fs';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Cli } from '../../../../cli/Cli';
import { CommandInfo } from '../../../../cli/CommandInfo';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./app-publish');

describe(commands.APP_PUBLISH, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const appResponse = {
    id: "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
    externalId: "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
    displayName: "Test App",
    distributionMethod: "organization"
  };


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
    loggerLogSpy = sinon.spy(logger, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      fs.readFileSync,
      fs.existsSync
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
    assert.strictEqual(command.name.startsWith(commands.APP_PUBLISH), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the filePath does not exist', async () => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = await command.validate({
      options: { filePath: 'invalid.zip' }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the filePath points to a directory', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({
      options: { filePath: './' }
    }, commandInfo);
    sinonUtil.restore([
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
  });

  it('validates for a correct input.', async () => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = await command.validate({
      options: {
        filePath: 'teamsapp.zip'
      }
    }, commandInfo);
    sinonUtil.restore([
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
  });

  it('adds new Teams app to the tenant app catalog', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return appResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, { options: { filePath: 'teamsapp.zip' } });
    assert(loggerLogSpy.calledWith(appResponse));
  });

  it('adds new Teams app to the tenant app catalog (debug)', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return appResponse;
      }

      throw 'Invalid request';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip' } });
    assert(loggerLogSpy.calledWith(appResponse));
  });

  it('correctly handles error when publishing an app', async () => {
    sinon.stub(request, 'post').callsFake(() => {
      throw 'An error has occurred';
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    await assert.rejects(command.action(logger, { options: { filePath: 'teamsapp.zip' } } as any), new CommandError('An error has occurred'));
  });
});
