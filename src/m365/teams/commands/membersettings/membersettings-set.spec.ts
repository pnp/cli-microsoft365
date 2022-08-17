import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./membersettings-set');

describe(commands.MEMBERSETTINGS_SET, () => {
  let log: string[];
  let logger: Logger;
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
    assert.strictEqual(command.name.startsWith(commands.MEMBERSETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        allowAddRemoveApps: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('sets the allowAddRemoveApps setting to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          memberSettings: {
            allowAddRemoveApps: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'true' }
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

  it('sets allowCreateUpdateChannels, allowCreateUpdateRemoveConnectors and allowDeleteChannels to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          memberSettings: {
            allowCreateUpdateChannels: true,
            allowCreateUpdateRemoveConnectors: true,
            allowDeleteChannels: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'true', allowCreateUpdateRemoveConnectors: 'true', allowDeleteChannels: 'true' }
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

  it('sets allowCreateUpdateChannels, allowCreateUpdateRemoveTabs and allowDeleteChannels to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.data) === JSON.stringify({
          memberSettings: {
            allowCreateUpdateChannels: false,
            allowCreateUpdateRemoveTabs: false,
            allowDeleteChannels: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'false', allowCreateUpdateRemoveTabs: 'false', allowDeleteChannels: 'false' }
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

  it('correctly handles error when updating member settings', (done) => {
    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.reject('An error has occurred');
    });

    command.action(logger, {
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'true' }
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

  it('fails validation if the teamId is not a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: 'invalid' } }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', async () => {
    const actual = await command.validate({ options: { teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('fails validation if allowAddRemoveApps is not a valid boolean', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateChannels is not a valid boolean', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveConnectors is not a valid boolean', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveTabs is not a valid boolean', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowDeleteChannels is not a valid boolean', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'invalid'
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if allowAddRemoveApps is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowAddRemoveApps is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is false', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'false'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is true', async () => {
    const actual = await command.validate({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'true'
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
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