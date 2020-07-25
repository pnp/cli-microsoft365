import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./o365group-renew');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import * as fs from 'fs';

describe(commands.O365GROUP_RENEW, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
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
  });

  afterEach(() => {
    Utils.restore([
      request.post,
      global.setTimeout
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      fs.readFileSync,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.O365GROUP_RENEW), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('renews expiration the specified group', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/renew/') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('renews expiration the specified group (debug)', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === 'https://graph.microsoft.com/v1.0/groups/28beab62-7540-4db1-a23f-29a6018a3848/renew/') {
        return Promise.resolve();
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, () => {
      try {
        assert(cmdInstanceLogSpy.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error when group is not found', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      return Promise.reject({ error: { 'odata.error': { message: { value: 'The remote server returned an error: (404) Not Found.' } } } });
    });

    cmdInstance.action({ options: { debug: false, id: '28beab62-7540-4db1-a23f-29a6018a3848' } }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The remote server returned an error: (404) Not Found.')));
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

  it('supports specifying id', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option.indexOf('--id') > -1) {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('fails validation if the id is not a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'abc' } });
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when the id is a valid GUID', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: '2c1ba4c4-cd9b-4417-832f-92a34bc34b2a' } });
    assert.strictEqual(actual, true);
  });
});