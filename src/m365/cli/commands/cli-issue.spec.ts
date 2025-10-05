import assert from 'assert';
import sinon from 'sinon';
import { cli } from '../../../cli/cli.js';
import { CommandInfo } from '../../../cli/CommandInfo.js';
import { Logger } from '../../../cli/Logger.js';
import { telemetry } from '../../../telemetry.js';
import { pid } from '../../../utils/pid.js';
import { session } from '../../../utils/session.js';
import commands from '../commands.js';
import { browserUtil } from '../../../utils/browserUtil.js';
import command from './cli-issue.js';

describe(commands.ISSUE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let openStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    (command as any).open = () => { };
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    openStub = sinon.stub(browserUtil, 'open').callsFake(async () => { return; });
  });

  afterEach(() => {
    openStub.restore();
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ISSUE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('accepts Bug issue Type', async () => {
    const actual = await command.validate({ options: { type: 'bug' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Command issue Type', async () => {
    const actual = await command.validate({ options: { type: 'command' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('accepts Sample issue Type', async () => {
    const actual = await command.validate({ options: { type: 'sample' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('rejects invalid issue type', async () => {
    const type = 'foo';
    const actual = await command.validate({ options: { type: type } }, commandInfo);
    assert.strictEqual(actual, `${type} is not a valid Issue type. Allowed values are bug, command, sample`);
  });

  it('Opens URL for a command (debug)', async () => {
    const commandUrl = 'https://aka.ms/cli-m365/new-command';

    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === commandUrl) {
        return;
      }
      throw 'Invalid url';
    });
    await command.action(logger, {
      options: {
        debug: true,
        type: 'command'
      }
    } as any);
    openStub.calledWith(commandUrl);
  });

  it('Opens URL for a bug (debug)', async () => {
    const bugUrl = 'https://aka.ms/cli-m365/bug';
    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === bugUrl) {
        return;
      }
      throw 'Invalid url';
    });
    await command.action(logger, {
      options: {
        debug: true,
        type: 'bug'
      }
    } as any);
    openStub.calledWith(bugUrl);
  });

  it('Opens URL for a sample (debug)', async () => {
    const sampleScriptUrl = 'https://aka.ms/cli-m365/new-sample-script';
    openStub.restore();
    openStub = sinon.stub(browserUtil, 'open').callsFake(async (url) => {
      if (url === sampleScriptUrl) {
        return;
      }
      throw 'Invalid url';
    });
    await command.action(logger, {
      options: {
        debug: true,
        type: 'sample'
      }
    } as any);
    openStub.calledWith(sampleScriptUrl);
  });
});
