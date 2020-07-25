import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./funsettings-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_FUNSETTINGS_LIST, () => {
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
      request.get
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_FUNSETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists fun settings of a Microsoft Teams team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315?$select=funSettings`) {
        return Promise.resolve({
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists fun settings of a Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/02bd9fd6-8f93-4758-87c3-1fb73740a315?$select=funSettings`) {
        return Promise.resolve({
          "funSettings": {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315"
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith(
          {
            "allowGiphy": true,
            "giphyContentRating": "moderate",
            "allowStickersAndMemes": true,
            "allowCustomMemes": false
          }
        ));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when retrieving funsettings', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        teamId: "02bd9fd6-8f93-4758-87c3-1fb73740a315"
      }
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

  it('fails validation if teamId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'invalid' }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when teamId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({
      options: { teamId: 'b1cf424e-f4f6-40b2-974e-6041524f4d66' }
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