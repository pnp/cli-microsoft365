import * as assert from 'assert';
import * as sinon from 'sinon';
import commands from '../commands';
import { Logger } from '../../../cli/Logger';
import { sinonUtil } from '../../../utils/sinonUtil';
import appInsights from '../../../appInsights';
import Command from '../../../Command';
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
      appInsights.trackEvent
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
    assert.deepStrictEqual((command as any).saveContextInfo(), undefined);
  });

});