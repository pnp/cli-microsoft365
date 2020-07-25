import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./messagingsettings-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_MESSAGINGSETTINGS_LIST, () => {
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
    assert.strictEqual(command.name.startsWith(commands.TEAMS_MESSAGINGSETTINGS_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists messaging settings for a Microsoft Team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "allowUserEditMessages": true,
          "allowUserDeleteMessages": true,
          "allowOwnerDeleteMessages": true,
          "allowTeamMentions": true,
          "allowChannelMentions": true
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists messaging settings for a Microsoft Team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: true } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "allowUserEditMessages": true,
          "allowUserDeleteMessages": true,
          "allowOwnerDeleteMessages": true,
          "allowTeamMentions": true,
          "allowChannelMentions": true
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when listing messaging settings', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject('An error has occurred');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", debug: false } }, (err?: any) => {
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
      options: {
        debug: false,
        teamId: 'invalid'
      }
    });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when a valid teamId is specified', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        debug: false,
        teamId: '2609af39-7775-4f94-a3dc-0dd67657e900'
      }
    });
    assert.strictEqual(actual, true);
  });

  it('lists all properties for output json', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/2609af39-7775-4f94-a3dc-0dd67657e900?$select=messagingSettings`) {
        return Promise.resolve({
          "messagingSettings": {
            "allowUserEditMessages": true,
            "allowUserDeleteMessages": true,
            "allowOwnerDeleteMessages": true,
            "allowTeamMentions": true,
            "allowChannelMentions": true
          }
        });
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({ options: { teamId: "2609af39-7775-4f94-a3dc-0dd67657e900", output: 'json', debug: false } }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith({
          "allowUserEditMessages": true,
          "allowUserDeleteMessages": true,
          "allowOwnerDeleteMessages": true,
          "allowTeamMentions": true,
          "allowChannelMentions": true
        }));
        done();
      }
      catch (e) {
        done(e);
      }
    });
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