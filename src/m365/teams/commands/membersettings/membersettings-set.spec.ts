import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./membersettings-set');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as chalk from 'chalk';

describe(commands.TEAMS_MEMBERSETTINGS_SET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
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
    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_MEMBERSETTINGS_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        name: 'Architecture',
        description: 'Architecture meeting'
      }
    });
    assert.strictEqual(actual, true);
    done();
  });

  it('sets the allowAddRemoveApps setting to true', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowAddRemoveApps: true
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'true' }
    }, (err?: any) => {
      try {
        assert.strictEqual(typeof err, 'undefined');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sets the allowAddRemoveApps setting to false', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/6703ac8a-c49b-4fd4-8223-28f0ac3a6402` &&
        JSON.stringify(opts.body) === JSON.stringify({
          memberSettings: {
            allowAddRemoveApps: false
          }
        })) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: true, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'false' }
    }, (err?: any) => {
      try {
        assert(cmdInstanceLogSpy.calledWith(chalk.green('DONE')));
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
        JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'true', allowCreateUpdateRemoveConnectors: 'true', allowDeleteChannels: 'true' }
    }, (err?: any) => {
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
        JSON.stringify(opts.body) === JSON.stringify({
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowCreateUpdateChannels: 'false', allowCreateUpdateRemoveTabs: 'false', allowDeleteChannels: 'false' }
    }, (err?: any) => {
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
    sinon.stub(request, 'patch').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: { debug: false, teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402', allowAddRemoveApps: 'true' }
    }, (err?: any) => {
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
    const actual = (command.validate() as CommandValidate)({ options: { teamId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55' } });
    assert.strictEqual(actual, true);
  });

  it('fails validation if allowAddRemoveApps is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateChannels is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveConnectors is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowCreateUpdateRemoveTabs is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('fails validation if allowDeleteChannels is not a valid boolean', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if allowAddRemoveApps is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowAddRemoveApps is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowAddRemoveApps: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateChannels is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateChannels: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveConnectors is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveConnectors: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowCreateUpdateRemoveTabs is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowCreateUpdateRemoveTabs: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is false', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'false'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('passes validation if allowDeleteChannels is true', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6f6fd3f7-9ba5-4488-bbe6-a789004d0d55',
        allowDeleteChannels: 'true'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});