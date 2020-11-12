import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../appInsights';
import auth from '../../../Auth';
import { Logger } from '../../../cli';
import Command, { CommandError } from '../../../Command';
import request from '../../../request';
import Utils from '../../../Utils';
import commands from '../commands';
const command: Command = require('./yammer-search');

describe(commands.YAMMER_SEARCH, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const searchResults: any = {
      "count": {
        "messages":  4,
        "groups":  2,
        "topics":  5,
        "users":  4
      },
      "messages": {
        "messages": [{
          "id": 11111
        },
        {
          "id": 11112
        },
        {
          "id": 11113
        },
        {
          "id": 11114
        }]
      },
      "groups": [
        {
          "id": 2221
        },
        {
          "id": 2222
        }
      ],
      "topics": [
        { 
          "id": 3331
        },
        { 
          "id": 3332
        },
        { 
          "id": 3333
        },
        { 
          "id": 3334
        },
        { 
          "id": 3335
        }
      ],
      "users": [
        { 
          "id": 4441
        },
        { 
          "id": 4442
        },
        { 
          "id": 4443
        },
        { 
          "id": 4444
        }
      ]        
    };

    const searchResults2: any = {
      "count": {
        "messages":  4,
        "groups":  2,
        "topics":  5,
        "users":  4
      },
      "messages": {
        "messages": []
      },
      "groups": [],
      "topics": [],
      "users": []        
    };

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: (msg: string) => {
        log.push(msg);
      },
      logRaw: (msg: string) => {
        log.push(msg);
      },
      logToStderr: (msg: string) => {
        log.push(msg);
      }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_SEARCH), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      return Promise.reject({
        "error": {
          "base": "An error has occurred."
        }
      });
    });

    command.action(logger, { options: { debug: false } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('does not pass validation without parameters', () => {
    const actual = command.validate({ options: {} });
    assert.notStrictEqual(actual, false);
  });

  it('passes validation with one parameter', () => {
    const actual = command.validate({ options: { search: '123123' } });
    assert.notStrictEqual(actual, false);
  });

  it('passes validation with parameters', () => {
    const actual = command.validate({ options: { search: '123', limit: 10, output: 'json' } });
    assert.strictEqual(actual, true);
  });

  it('limit can only be used with output json', () => {
    const actual = command.validate({ options: { search: '123', limit: 10 } });
    assert.notStrictEqual(actual, true);
  });

  it('limit must be a number', () => {
    const actual = command.validate({ options: { search: '123', limit: 'abc', output: 'json' } });
    assert.notStrictEqual(actual, true);
  });

  it('query must be a string', () => {
    const actual = command.validate({ options: { search: 'abc' } });
    assert.strictEqual(actual, true);
  });

  it('query must be a string', () => {
    const actual = command.validate({ options: { search: 123 } });
    assert.notStrictEqual(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options();
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('returns the summary', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=1') {
        return Promise.resolve(searchResults);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { search:"contents" } } as any, (err?: any) => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].messages, 4)
        assert.strictEqual(loggerLogSpy.lastCall.args[0].groups, 2)
        assert.strictEqual(loggerLogSpy.lastCall.args[0].topics, 5)
        assert.strictEqual(loggerLogSpy.lastCall.args[0].users, 4)
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns limited results', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=1') {
        return Promise.resolve(searchResults);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { search:"contents", limit: 1, output: "json"  } } as any, (err?: any) => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.messages, 4, "summary returns 4 messages")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.groups, 2, "sumamry returns 2 groups")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.topics, 5, "summary return two topics")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.users, 4, "summary returns 4 users")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].messages.length, 1, "message arary returns 1 message")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].groups.length, 1, "groups array returns 1 group")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].topics.length, 1, "topics array returns 1 topic")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].users.length, 1, "users array returns 1 user")
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all results', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=1') {
        return Promise.resolve(searchResults);
      }
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=2') {
        return Promise.resolve(searchResults2);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { search:"contents", output: "json" } } as any, (err?: any) => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.messages, 4, "summary returns 4 messages")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.groups, 2, "sumamry returns 2 groups")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.topics, 5, "summary return two topics")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].summary.users, 4, "summary returns 4 users")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].messages.length, 4, "message arary returns 4 entries")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].groups.length, 2, "groups array returns 2 groups")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].topics.length, 5, "topics array returns 2 topics")
        assert.strictEqual(loggerLogSpy.lastCall.args[0].users.length, 4, "users array returns 4 users")
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error in loop', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=1') {
        return Promise.resolve(searchResults);
      }
      if (opts.url === 'https://www.yammer.com/api/v1/search.json?search=contents&page=2') {
        return Promise.reject({
          "error": {
            "base": "An error has occurred."
          }
        });
      }

      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { search:"contents", output: "json" } } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError("An error has occurred.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
}); 