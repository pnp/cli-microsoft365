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
import command, { options } from './cli-issue.js';

describe(commands.ISSUE, () => {
  let log: any[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  let openStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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

  it('fails validation with no options', () => {
    const actual = commandOptionsSchema.safeParse({});
    assert.notStrictEqual(actual.success, true);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'bug', unknownOption: 'value' });
    assert.notStrictEqual(actual.success, true);
  });

  it('accepts Bug issue Type', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'bug' });
    assert.strictEqual(actual.success, true);
  });

  it('accepts Command issue Type', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'command' });
    assert.strictEqual(actual.success, true);
  });

  it('accepts Sample issue Type', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'sample' });
    assert.strictEqual(actual.success, true);
  });

  it('rejects invalid issue type', () => {
    const actual = commandOptionsSchema.safeParse({ type: 'foo' });
    assert.strictEqual(actual.success, false);
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
      options: commandOptionsSchema.parse({
        debug: true,
        type: 'command'
      })
    });
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
      options: commandOptionsSchema.parse({
        debug: true,
        type: 'bug'
      })
    });
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
      options: commandOptionsSchema.parse({
        debug: true,
        type: 'sample'
      })
    });
    openStub.calledWith(sampleScriptUrl);
  });
});
