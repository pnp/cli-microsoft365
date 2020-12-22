import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Logger } from '../../../cli';
import Command from '../../../Command';
import Utils from '../../../Utils';
import commands from '../commands';
const command: Command = require('./cli-reconsent');

describe(commands.COMPLETION_SH_SETUP, () => {
  let log: string[];
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

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.RECONSENT), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('generates file with commands info', (done) => {
    command.action(logger, { options: { debug: false } }, () => {
      try {
        assert(log[0].indexOf('/oauth2/authorize?client_id') > -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});