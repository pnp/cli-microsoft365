import commands from '../../commands';
import Command, { CommandOption, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-teamify');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.O365GROUP_TEAMIFY, () => {
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
    assert.equal(command.name.startsWith(commands.O365GROUP_TEAMIFY), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('o365group timify success', (done) => {
    const requestStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: false, groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    }, () => {
      try {
        assert.equal(requestStub.lastCall.args[0].url, 'https://graph.microsoft.com/beta/teams');
        assert.equal(requestStub.lastCall.args[0].body["template@odata.bind"], 'https://graph.microsoft.com/beta/teamsTemplates(\'standard\')');
        assert.equal(requestStub.lastCall.args[0].body["group@odata.bind"], `https://graph.microsoft.com/v1.0/groups('8231f9f2-701f-4c6e-93ce-ecb563e3c1ee')`);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('o365group timify success (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/teams`) {
        return Promise.resolve({});
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: { debug: true, groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    }, () => {
      try {
        assert.notEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('should handle Microsoft graph error response', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
     
        return Promise.reject({
          "error": {
            "code": "ItemNotFound",
            "message": "Error: Failed to execute Templates backend request CreateTeamFromGroupWithTemplateRequest.",
            "innerError": {
              "request-id": "27b49647-a335-48f8-9a7c-f1ed9b976aaa",
              "date": "2019-04-05T12:16:48"
            }
          }
        });
    });

    cmdInstance.action({
      options: { debug: false, groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' }
    }, (err?: any) => {
      try {
        assert.equal(err.message, 'Error: Failed to execute Templates backend request CreateTeamFromGroupWithTemplateRequest.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the groupId is not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('fails validation if the groupId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation if the groupId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } });
    assert.equal(actual, true);
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
    assert(find.calledWith(commands.O365GROUP_TEAMIFY));
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