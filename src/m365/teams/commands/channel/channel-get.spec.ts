import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./channel-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.TEAMS_CHANNEL_GET, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  
  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
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
      vorpal.find,
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
    assert.equal(command.name.startsWith(commands.TEAMS_CHANNEL_GET), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('fails validation if the teamId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'invalid',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    });
    assert.notEqual(actual, true);
  });

  it('fails validation when no channelId is specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('correctly validates the when all options are valid', () => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    });
    assert.equal(actual, true);
  });

  it('fails to get channel information due to wrong channel id', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "Failed to execute Skype backend request GetThreadS2SRequest.",
            "innerError": {
              "request-id": "4bebd0d2-d154-491b-b73f-d59ad39646fb",
              "date": "2019-04-06T13:40:51"
            }
          }
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError(`Failed to execute Skype backend request GetThreadS2SRequest.`)));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel information for the Microsoft Teams team', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "channel1",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = cmdInstanceLogSpy.lastCall;
        assert.equal(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
        assert.equal(call.args[0].displayName, 'channel1');
        assert.equal(call.args[0].description, null);
        assert.equal(call.args[0].email, '');
        assert.equal(call.args[0].webUrl, 'https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should get channel information for the Microsoft Teams team (debug)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/teams/39958f28-eefb-4006-8f83-13b6ac2a4a7f/channels/19%3A493665404ebd4a18adb8a980a31b4986%40thread.skype`) {
        return Promise.resolve({
          "id": "19:493665404ebd4a18adb8a980a31b4986@thread.skype",
          "displayName": "channel1",
          "description": null,
          "email": "",
          "webUrl": "https://teams.microsoft.com/l/channel/19%3a493665404ebd4a18adb8a980a31b4986%40thread.skype/channel1?groupId=39958f28-eefb-4006-8f83-13b6ac2a4a7f&tenantId=ea1787c6-7ce2-4e71-be47-5e0deb30f9e4"
        });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: true,
        teamId: '39958f28-eefb-4006-8f83-13b6ac2a4a7f',
        channelId: '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
      }
    }, () => {
      try {
        const call: sinon.SinonSpyCall = cmdInstanceLogSpy.getCall(cmdInstanceLogSpy.callCount - 2);
        assert.equal(call.args[0].id, '19:493665404ebd4a18adb8a980a31b4986@thread.skype');
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

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.TEAMS_CHANNEL_GET));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});