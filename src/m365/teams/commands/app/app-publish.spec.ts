import * as assert from 'assert';
import * as chalk from 'chalk';
import * as fs from 'fs';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./app-publish');

describe(commands.TEAMS_APP_PUBLISH, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
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
    loggerLogSpy = sinon.spy(logger, 'log');
    loggerLogToStderrSpy = sinon.spy(logger, 'logToStderr');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      fs.readFileSync,
      fs.existsSync
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAMS_APP_PUBLISH), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the filePath does not exist', (done) => {
    sinon.stub(fs, 'existsSync').callsFake(() => false);
    const actual = command.validate({
      options: { filePath: 'invalid.zip' }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails validation if the filePath points to a directory', (done) => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => true);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = command.validate({
      options: { filePath: './' }
    });
    Utils.restore([
      fs.lstatSync
    ]);
    assert.notStrictEqual(actual, true);
    done();
  });

  it('validates for a correct input.', (done) => {
    const stats: fs.Stats = new fs.Stats();
    sinon.stub(stats, 'isDirectory').callsFake(() => false);
    sinon.stub(fs, 'existsSync').callsFake(() => true);
    sinon.stub(fs, 'lstatSync').callsFake(() => stats);

    const actual = command.validate({
      options: {
        filePath: 'teamsapp.zip'
      }
    });
    Utils.restore([
      fs.lstatSync
    ]);
    assert.strictEqual(actual, true);
    done();
  });

  it('adds new Teams app to the tenant app catalog', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return Promise.resolve({
          "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
          "externalId": "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
          "name": "Test App",
          "version": "1.0.0",
          "distributionMethod": "organization"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: false, filePath: 'teamsapp.zip' } }, () => {
      try {
        assert(loggerLogSpy.calledWith("e3e29acb-8c79-412b-b746-e6c39ff4cd22"));
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('adds new Teams app to the tenant app catalog (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps`) {
        return Promise.resolve({
          "id": "e3e29acb-8c79-412b-b746-e6c39ff4cd22",
          "externalId": "b5561ec9-8cab-4aa3-8aa2-d8d7172e4311",
          "name": "Test App",
          "version": "1.0.0",
          "distributionMethod": "organization"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: true, filePath: 'teamsapp.zip' } }, () => {
      try {
        assert(loggerLogSpy.calledWith("e3e29acb-8c79-412b-b746-e6c39ff4cd22"));
        assert(loggerLogToStderrSpy.calledWith(chalk.green('DONE')));

        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when publishing an app', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    sinon.stub(fs, 'readFileSync').callsFake(() => '123');

    command.action(logger, { options: { debug: false, filePath: 'teamsapp.zip' } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      } catch (e) {
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