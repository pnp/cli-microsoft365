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
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
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
    assert.equal(command.name.startsWith(commands.YAMMER_NETWORK_LIST), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
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
        assert.equal(cmdInstanceLogSpy.lastCall.args[0][0].id, '123')
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
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
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
        assert.equal(cmdInstanceLogSpy.lastCall.args[0][0].id, '123');
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
        assert.equal(cmdInstanceLogSpy.lastCall.args[0][0].id, '123')
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation without parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.equal(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { includeSuspended: true } });
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
    assert(find.calledWith(commands.YAMMER_NETWORK_LIST));
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