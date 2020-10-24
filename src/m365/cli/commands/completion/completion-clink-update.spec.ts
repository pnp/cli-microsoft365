import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import { autocomplete } from '../../../../autocomplete';
import { Logger } from '../../../../cli';
import Command from '../../../../Command';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./completion-clink-update');

describe(commands.COMPLETION_CLINK_UPDATE, () => {
  let log: string[];
  let logger: Logger;
  let generateClinkCompletionStub: sinon.SinonStub;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    Utils.restore([
      appInsights.trackEvent,
      autocomplete.getClinkCompletion
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.COMPLETION_CLINK_UPDATE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('builds command completion', (done) => {
    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(generateClinkCompletionStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});