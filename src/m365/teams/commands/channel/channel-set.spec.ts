import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./channel-set');

describe(commands.TEAMS_CHANNEL_SET, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      request.get,
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_CHANNEL_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly validates the arguments', () => {
    const actual = command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelName: 'Reviews',
        newChannelName: 'Gen',
        description: 'this is a new description'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = command.validate({
      options: {
        teamId: 'invalid',
        channelName: 'Reviews',
        newChannelName: 'Gen',
        description: 'this is a new description'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation when channelName is General', (done) => {
    const actual = command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelName: 'General',
        newChannelName: 'Reviews',
        description: 'this is a new description'
      }
    });
    assert.notStrictEqual(actual, true);
    done();
  });

  it('fails to patch channel updates for the Microsoft Teams team when channel does not exists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'Latest'`) > -1) {
        return Promise.resolve({ value: [] });
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: true,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Latest',
        newChannelName: 'New Review',
        description: 'New Review'
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError(`The specified channel does not exist in the Microsoft Teams team`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly patches channel updates for the Microsoft Teams team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if ((opts.url as string).indexOf(`channels?$filter=displayName eq 'Review'`) > -1) {
        return Promise.resolve({
          value:
            [
              {
                "id": "19:8a53185a51ac44a3aef27397c3dfebfc@thread.skype",
                "displayName": "Review",
                "description": "Updated by CLI"
              }]
        });
      }
      return Promise.reject('Invalid request');
    });
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (((opts.url as string).indexOf(`channels/19:8a53185a51ac44a3aef27397c3dfebfc@thread.skype`) > -1) &&
        JSON.stringify(opts.data) === JSON.stringify({ displayName: "New Review", description: "New Review" })
      ) {
        return Promise.resolve({});
      }
      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        teamId: '00000000-0000-0000-0000-000000000000',
        channelName: 'Review',
        newChannelName: 'New Review',
        description: 'New Review'
      }
    } as any, (err?: any) => {
      try {
        assert(loggerLogSpy.notCalled);
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