import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./message-get');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.YAMMER_MESSAGE_GET, () => {
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;
  let firstMessage: any = {"sender_id":1496550646, "replied_to_id":1496550647,"id":10123190123123,"thread_id": "", group_id: 11231123123, created_at: "2019/09/09 07:53:18 +0000", "content_excerpt": "message1"};
  let secondMessage: any = {"sender_id":1496550640, "replied_to_id":"","id":10123190123124,"thread_id": "", group_id: "", created_at: "2019/09/08 07:53:18 +0000", "content_excerpt": "message2"};

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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_MESSAGE_GET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('id must be a number', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 'nonumber' } });
    assert.notStrictEqual(actual, true);
  });

  it('id is required', () => {
    const actual = (command.validate() as CommandValidate)({ options: { } });
    assert.notStrictEqual(actual, true);
  });

  it('calls the messaging endpoint with the right parameters', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123123.json') {
        return Promise.resolve(firstMessage);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { id:10123190123123, debug: true } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].id, 10123190123123)
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

  it('calls the messaging endpoint with id and json and json', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/messages/10123190123124.json') {
        return Promise.resolve(secondMessage);
      }
      return Promise.reject('Invalid request');
    });
    cmdInstance.action({ options: { debug: true, id:10123190123124, output: "json" } }, (err?: any) => {
      try {
        assert.strictEqual(cmdInstanceLogSpy.lastCall.args[0].id, 10123190123124);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('passes validation with parameters', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 10123123 }});
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