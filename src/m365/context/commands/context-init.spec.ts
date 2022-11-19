import * as assert from 'assert';
import * as sinon from 'sinon';
import commands from '../commands';
import { Logger } from '../../../cli/Logger';
import { sinonUtil } from '../../../utils/sinonUtil';
import appInsights from '../../../appInsights';
import Command from '../../../Command';
import * as fs from 'fs';
const command: Command = require('./context-init');

describe(commands.INIT, () => {
  let log: any[];
  let logger: Logger;

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
  });

  afterEach(() => {
    sinonUtil.restore([
      appInsights.trackEvent,
      fs.existsSync,
      fs.readFileSync
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

  it('writes a context to the .m365rc.json file', async () => {
    await command.action(logger, { options: { verbose: true } });

    const fileContent = fs.readFileSync('.m365rc.json', 'utf8');
    const contextInfo = JSON.parse(fileContent);

    assert.deepStrictEqual(Object.keys(contextInfo).indexOf('context') > -1, true);
  });

});