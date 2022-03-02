import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./guestsettings-set');

describe(commands.GUESTSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;

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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.patch
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
    assert.strictEqual(command.name.startsWith(commands.GUESTSETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('sets the allowDeleteChannels setting to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          guestSettings: {
            allowDeleteChannels: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: 'true' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets allowCreateUpdateChannels and allowDeleteChannels to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          guestSettings: {
            allowCreateUpdateChannels: true,
            allowDeleteChannels: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'true', allowDeleteChannels: 'true' }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when updating guest settings', (done) => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowDeleteChannels: 'true' }
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

  it('fails validation if the teamId is not a valid GUID', () => {
    const actual = command.validate({ options: { teamId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = command.validate({ options: { teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if allowDeleteChannels is not a valid boolean', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateChannels is not a valid boolean', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is false', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is true', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', () => {
    const actual = command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'true'
      }
    });
    assert.strictEqual(actual, true);
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