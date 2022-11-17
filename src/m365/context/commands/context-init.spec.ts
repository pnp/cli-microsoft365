import * as assert from 'assert';
import * as sinon from 'sinon';
import * as fs from 'fs';
import commands from '../commands';
import { Logger } from '../../../cli/Logger';
import { sinonUtil } from '../../../utils/sinonUtil';
import appInsights from '../../../appInsights';
import Command from '../../../Command';
import * as ContextCommand from '../../base/ContextCommand';
import { Hash } from '../../../utils/types';
const command: Command = require('./context-init');

describe(commands.INIT, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  const contextInfo: Hash = {};

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
  });

  afterEach(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync,
      fs.writeFileSync
    ]);
  });

  after(() => {
    sinonUtil.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.INIT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('retrieves a context', async () => {
    sinon.stub(ContextCommand, 'default').callsFake(async () => { });

    await command.action(logger, { options: { verbose: true } });
    const test = loggerLogSpy.lastCall;
    assert.strictEqual(Object.keys(test.args[0]).length, Object.keys(contextInfo).length);
  });

});