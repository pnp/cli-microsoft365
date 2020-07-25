import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./network-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_NETWORK_LIST, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_NETWORK_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('calls the networking endpoint without parameter', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
            {
              "id": 123,
              "name": "Network1",
              "email": "email@mail.com",
              "community": true,
              "permalink": "network1-link",
              "web_url": "https://www.yammer.com/network1-link"
            },
            {
              "id": 456,
              "name": "Network2",
              "email": "email2@mail.com",
              "community": false,
              "permalink": "network2-link",
              "web_url": "https://www.yammer.com/network2-link"
            }
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 123)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

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

  it('calls the networking endpoint without parameter and json', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
            {
              "id": 123,
              "name": "Network1",
              "email": "email@mail.com",
              "community": true,
              "permalink": "network1-link",
              "web_url": "https://www.yammer.com/network1-link"
            },
            {
              "id": 456,
              "name": "Network2",
              "email": "email2@mail.com",
              "community": false,
              "permalink": "network2-link",
              "web_url": "https://www.yammer.com/network2-link"
            }
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, output: "json" } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 123);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('calls the networking endpoint with parameter', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/networks/current.json') {
        return Promise.resolve(
          [
            {
              "id": 123,
              "name": "Network1",
              "email": "email@mail.com",
              "community": true,
              "permalink": "network1-link",
              "web_url": "https://www.yammer.com/network1-link"
            },
            {
              "id": 456,
              "name": "Network2",
              "email": "email2@mail.com",
              "community": false,
              "permalink": "network2-link",
              "web_url": "https://www.yammer.com/network2-link"
            }
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, includeSuspended: true } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0][0].id, 123);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation without parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { includeSuspended: true } });
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