import commands from '../commands';
import Command from '../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
const command: Command = require('./cli-reconsent');
import * as assert from 'assert';
import Utils from '../../../Utils';

describe(commands.COMPLETION_SH_SETUP, () => {
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
  });

  beforeEach(() => {
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
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
    cmdInstance.action({ options: { debug: false } }, () => {
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