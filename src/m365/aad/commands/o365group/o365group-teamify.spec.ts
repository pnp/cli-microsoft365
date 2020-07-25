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
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_TEAMIFY), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('validates for a correct input.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee'
      }
    });
    assert.strictEqual(actual, true);
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
        assert.strictEqual(requestStub.lastCall.args[0].url, 'https://graph.microsoft.com/beta/teams');
        assert.strictEqual(requestStub.lastCall.args[0].body["template@odata.bind"], 'https://graph.microsoft.com/beta/teamsTemplates(\'standard\')');
        assert.strictEqual(requestStub.lastCall.args[0].body["group@odata.bind"], `https://graph.microsoft.com/v1.0/groups('8231f9f2-701f-4c6e-93ce-ecb563e3c1ee')`);
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
        assert.notStrictEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
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
        assert.strictEqual(err.message, 'Error: Failed to execute Templates backend request CreateTeamFromGroupWithTemplateRequest.');
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the groupId is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: 'invalid' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation if the groupId is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { groupId: '8231f9f2-701f-4c6e-93ce-ecb563e3c1ee' } });
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