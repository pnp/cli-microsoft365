import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import { Logger } from '../../../cli';
import Command from '../../../Command';
import Utils from '../../../Utils';
import commands from '../commands';
import Sinon = require('sinon');

const command: Command = require('./cli-issue');

describe(commands.ISSUE, () => {
  let log: any[];
  let logger: Logger;
  let openBrowserSpy: Sinon.SinonSpy;

  before(() => {
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    (command as any).open = () => { };
    openBrowserSpy = sinon.spy(command as any, 'openBrowser');
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
    openBrowserSpy.resetHistory();
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent
    ]);
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.ISSUE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('accepts Bug issue Type', () => {
    const actual = command.validate({ options: { type: 'bug' } });
    assert.strictEqual(actual, true);
  });

  it('accepts Command issue Type', () => {
    const actual = command.validate({ options: { type: 'command' } });
    assert.strictEqual(actual, true);
  });

  it('accepts Sample issue Type', () => {
    const actual = command.validate({ options: { type: 'sample' } });
    assert.strictEqual(actual, true);
  });

  it('rejects invalid issue type', () => {
    const type = 'foo';
    const actual = command.validate({ options: { type: type } });
    assert.strictEqual(actual, `${type} is not a valid Issue type. Allowed values are bug, command, sample`);
  });

  it('Opens URL for a command (debug)', (done) => {
    command.action(logger, {
      options: {
        debug: true,
        type: 'command'
      }
    } as any, () => {
      try {
        openBrowserSpy.calledWith("https://aka.ms/cli-m365/new-command");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Opens URL for a bug (debug)', (done) => {
    command.action(logger, {
      options: {
        debug: true,
        type: 'bug'
      }
    } as any, () => {
      try {
        openBrowserSpy.calledWith("https://aka.ms/cli-m365/bug");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Opens URL for a sample (debug)', (done) => {
    command.action(logger, {
      options: {
        debug: true,
        type: 'sample'
      }
    } as any, () => {
      try {
        openBrowserSpy.calledWith("https://aka.ms/cli-m365/new-sample-script");
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});
