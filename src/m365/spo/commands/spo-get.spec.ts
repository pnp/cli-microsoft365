import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Logger } from '../../../cli';
import Command, { CommandError } from '../../../Command';
import Utils from '../../../Utils';
import commands from '../commands';
const command: Command = require('./spo-get');

describe(commands.GET, () => {
  let log: string[];
  let logger: Logger;
  let loggerSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'storeConnectionInfo').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    loggerSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    auth.service.spoUrl = undefined;
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      auth.storeConnectionInfo,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('gets SPO URL when no URL was get previously', (done) => {
    auth.service.spoUrl = undefined;

    command.action(logger, {
      options: {
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(loggerSpy.calledWith({
          SpoUrl: ''
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('gets SPO URL when other URL was get previously', (done) => {
    auth.service.spoUrl = 'https://northwind.sharepoint.com';

    command.action(logger, {
      options: {
        output: 'json',
        debug: true
      }
    }, () => {
      try {
        assert(loggerSpy.calledWith({
          SpoUrl: 'https://northwind.sharepoint.com'
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('throws error when trying to get SPO URL when not logged in to O365', (done) => {
    auth.service.connected = false;

    command.action(logger, { options: {} } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("Log in to Microsoft 365 first")));
        assert.strictEqual(auth.service.spoUrl, undefined);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Contains the correct options', () => {
    const options = command.options();
    let containsOutputOption = false;
    let containsVerboseOption = false;
    let containsDebugOption = false;
    let containsQueryOption = false;

    options.forEach(o => {
      if (o.option.indexOf('--output') > -1) {
        containsOutputOption = true;
      } else if (o.option.indexOf('--verbose') > -1) {
        containsVerboseOption = true;
      } else if (o.option.indexOf('--debug') > -1) {
        containsDebugOption = true;
      } else if (o.option.indexOf('--query') > -1) {
        containsQueryOption = true;
      }
    });

    assert(options.length === 4, "Wrong amount of options returned");
    assert(containsOutputOption, "Output option not available");
    assert(containsVerboseOption, "Verbose option not available");
    assert(containsDebugOption, "Debug option not available");
    assert(containsQueryOption, "Query option not available");
  });

  it('passes validation without any extra options', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });
});