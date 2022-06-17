import * as assert from 'assert';
import chalk = require('chalk');
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./team-archive');

describe(commands.TEAM_ARCHIVE, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let loggerLogToStderrSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    sinonUtil.restore([
      request.get,
      request.post
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TEAM_ARCHIVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if the teamId is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        teamId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if the id is not a valid guid.', async () => {
    const actual = await command.validate({
      options: {
        id: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the input is correct', async () => {
    const actual = await command.validate({
      options: {
        teamId: '15d7a78e-fd77-4599-97a5-dbb6372846c5'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation when no option is specified', async () => {
    const actual = await command.validate({
      options: {
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when all options are specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Finance',
        id: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both id and name are specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Finance',
        id: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when both teamId and name are specified', async () => {
    const actual = await command.validate({
      options: {
        name: 'Finance',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('logs deprecation warning when option teamId is specified', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
      try {
        assert(loggerLogToStderrSpy.calledWith(chalk.yellow(`Option 'teamId' is deprecated. Please use 'id' instead.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails when team name does not exist', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Finance'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams",
          "@odata.count": 1,
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": []
            }
          ]
        }
        );
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        name: 'Finance',
        confirm: true
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified team does not exist in the Microsoft Teams`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('archives a Microsoft Team by id', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/f5dba91d-6494-4d5e-89a7-ad832f6946d6/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
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

  it('archives a Microsoft Team by name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'Finance'`) {
        return Promise.resolve({
          "value": [
            {
              "id": "00000000-0000-0000-0000-000000000000",
              "resourceProvisioningOptions": ["Team"]
            }
          ]
        });
      }
      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/00000000-0000-0000-0000-000000000000/archive`) {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        name: 'Finance'
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
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
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
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
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
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6'
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

  it('correctly handles error when archiving a team', (done) => {
    sinon.stub(request, 'post').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: {
        id: 'f5dba91d-6494-4d5e-89a7-ad832f6946d6',
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
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});