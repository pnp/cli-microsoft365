import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./team-archive');

describe(commands.TEAMS_TEAM_ARCHIVE, () => {
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
      request.post
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_TEAM_ARCHIVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = command.validate({
      options: {
        teamId: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', () => {
    const actual = command.validate({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('archives a Microsoft Team', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      }
    } as any, () => {
      try {
        assert(loggerLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('archives a Microsoft Teams teams when \'shouldSetSpoSiteReadOnlyForMembers\' specified', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
        shouldSetSpoSiteReadOnlyForMembers: true
      }
    } as any, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].data.shouldSetSpoSiteReadOnlyForMembers, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should set to false value when \'shouldSetSpoSiteReadOnlyForMembers\' not specified', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      }
    } as any, () => {
      try {
        assert.strictEqual(postStub.lastCall.args[0].data["shouldSetSpoSiteReadOnlyForMembers"], false);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should correctly handle graph error response', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.reject(
          {
              "error": {
                  "code": "ItemNotFound",
                  "message": "No team found with Group Id f5dba91d-6494-4d5e-89a7-ad832f6946d6",
                  "innerError": {
                      "request-id": "ad0c0a4f-a4fc-4567-8ae1-1150db48b620",
                      "date": "2019-04-05T15:51:43"
                  }
              }
          });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(err.message, 'No team found with Group Id f5dba91d-6494-4d5e-89a7-ad832f6946d6');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('archives a Microsoft Team (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
        debug: true
      }
    } as any, () => {
      try {
        assert(loggerLogToStderrSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when archiving a team', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        teamId: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
        debug: false
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
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