import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./guestsettings-list');

describe(commands.GUESTSETTINGS_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.GUESTSETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists guest settings for a Microsoft Team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return Promise.resolve({
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "allowCreateUpdateChannels": false,
          "allowDeleteChannels": false
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists guest settings for a Microsoft Team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return Promise.resolve({
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: true } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "allowCreateUpdateChannels": false,
          "allowDeleteChannels": false
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing guest settings for a Microsoft Team', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if teamId is not a valid GUID', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        teamId: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when teamId is valid', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        teamId: '2609af39-7775-4f94-a3dc-0dd67657e900'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=guestSettings`) {
        return Promise.resolve({
          "guestSettings": {
            "allowCreateUpdateChannels": false,
            "allowDeleteChannels": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, { options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", output: 'json', debug: false } }, () => {
      try {
        assert(loggerLogSpy.calledWith({
          "allowCreateUpdateChannels": false,
          "allowDeleteChannels": false
        }));
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