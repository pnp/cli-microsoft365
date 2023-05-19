import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli/Logger';
import Command from '../../../../Command';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import commands from '../../commands';
const command: Command = require('./completion-clink-update');

describe(commands.COMPLETION_CLINK_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let generateClinkCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    sinon.stub(session, 'getId').callsFake(() => '');
    generateClinkCompletionStub = sinon.stub(autocomplete, 'getClinkCompletion').callsFake(() => '');
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
    generateClinkCompletionStub.reset();
  });

  after(() => {
    sinon.restore();
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_CLINK_UPDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('builds command completion', async () => {
    await command.action(logger, { options: {} });
    assert(generateClinkCompletionStub.called);
  });
});
