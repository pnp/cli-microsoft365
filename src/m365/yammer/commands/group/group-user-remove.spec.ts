import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./group-user-remove');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_GROUP_USER_REMOVE, () => {
  let log: string[];
  let cmdInstance: any;
  let requests: any[];
  let promptOptions: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
      },
      prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
        promptOptions = options;
        cb({ continue: false });
      }
    };
    requests = [];
  });

  afterEach(() => {
    Utils.restore([
      request.delete
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_GROUP_USER_REMOVE), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'delete').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: false } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123 } });
    assert.strictEqual(actual, true);
  });

  it('id must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('userId must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10, userId: 'abc' } });
    assert.notStrictEqual(actual, true);
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

  it('calls the service if the current user is removed from the group', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: true, id: 1231231 } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service if the user 989998789 is removed from the group 1231231 with the confirm command', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1231231, userId: 989998789, confirm: true } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the service if the user 989998789 is removed from the group 1231231', (done) => {
    const requestDeleteStub = sinon.stub(request, 'delete').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/group_memberships.json') {
        return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: true });
    };

    cmdInstance.action({ options: { debug: true, id: 1231231, userId: 989998789 } }, () => {
      try {
        assert(requestDeleteStub.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('prompts before when confirmation argument not passed', (done) => {
    cmdInstance.action({ options: { debug: false, id: 1231231, userId: 989998789 } }, () => {
      let promptIssued = false;

      if (promptOptions && promptOptions.type === 'confirm') {
        promptIssued = true;
      }

      try {
        assert(promptIssued);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('aborts execution when prompt not confirmed', (done) => {
    cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
      cb({ continue: false });
    };
    cmdInstance.action({ options: { debug: false, id: 1231231, userId: 989998789 } }, () => {
      try {
        assert(requests.length === 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 