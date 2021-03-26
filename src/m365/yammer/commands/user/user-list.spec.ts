import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import Utils from '../../../../Utils';
import commands from '../../commands';
const command: Command = require('./user-list');

describe(commands.YAMMER_USER_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

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
    assert.strictEqual(command.name.startsWith(commands.YAMMER_USER_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'full_name', 'email']);
  });

  it('returns all network users', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: {} } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns all network users using json', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" }
          ]
        );
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { output: 'json' } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sorts network users by messages', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&sort_by=messages') {
        return Promise.resolve([
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        ]);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { sortBy: "messages" } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fakes the return of more results', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }],
          more_available: true
        });
      }
      else {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
            { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }],
          more_available: false
        });
      }
    });
    command.action(logger, { options: { output: 'json' } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 4);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fakes the return of more than 50 entries', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }]
        );
      }
      else {
        return Promise.resolve([
          { "type": "user", "id": 14965556, "network_id": 801445, "state": "active", "full_name": "Daniela Kiener" },
          { "type": "user", "id": 12310090123, "network_id": 801445, "state": "active", "full_name": "Carlo Lamber" }]);
      }
    });
    command.action(logger, { options: { output: 'debug' } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 52);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fakes the return of more results with exception', (done) => {
    let i: number = 0;

    sinon.stub(request, 'get').callsFake(() => {
      if (i++ === 0) {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }],
          more_available: true
        });
      }
      else {
        return Promise.reject({
          "error": {
            "base": "An error has occurred."
          }
        });
      }
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

  it('sorts network users by messages', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&sort_by=messages') {
        return Promise.resolve([
          { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
          { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }
        ]);
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { sortBy: "messages" } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sorts users in reverse order', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&reverse=true') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { reverse: true } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('sorts users in reverse order in a group and limits the user to 2', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1&reverse=true') {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" },
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550643, "network_id": 801445, "state": "active", "full_name": "Daniela Lamber" }],
          has_more: true
        });
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { groupId: 5785177, reverse: true, limit: 2 } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550647);
        assert.strictEqual(loggerLogSpy.lastCall.args[0].length, 2);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns users of a specific group', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users/in_group/5785177.json?page=1') {
        return Promise.resolve({
          users: [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" }, { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" }],
          has_more: false
        });
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { groupId: 5785177 } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('returns users starting with the letter P', function (done) {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === 'https://www.yammer.com/api/v1/users.json?page=1&letter=P') {
        return Promise.resolve(
          [
            { "type": "user", "id": 1496550646, "network_id": 801445, "state": "active", "full_name": "Patrick Lamber" },
            { "type": "user", "id": 1496550647, "network_id": 801445, "state": "active", "full_name": "Sergio Cappelletti" }]
        );
      }
      return Promise.reject('Invalid request');
    });
    command.action(logger, { options: { letter: "P" } } as any, () => {
      try {
        assert.strictEqual(loggerLogSpy.lastCall.args[0][0].id, 1496550646);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles error', (done) => {
    sinon.stub(request, 'get').callsFake(() => {
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

  it('passes validation without parameters', () => {
    const actual = command.validate({ options: {} });
    assert.strictEqual(actual, true);
  });

  it('passes validation with parameters', () => {
    const actual = command.validate({ options: { letter: "A" } });
    assert.strictEqual(actual, true);
  });

  it('letter does not allow numbers', () => {
    const actual = command.validate({ options: { letter: "1" } });
    assert.notStrictEqual(actual, true);
  });

  it('groupId must be a number', () => {
    const actual = command.validate({ options: { groupId: "aasdf" } });
    assert.notStrictEqual(actual, true);
  });

  it('limit must be a number', () => {
    const actual = command.validate({ options: { limit: "aasdf" } });
    assert.notStrictEqual(actual, true);
  });

  it('sortBy validation check', () => {
    const actual = command.validate({ options: { sortBy: "aasdf" } });
    assert.notStrictEqual(actual, true);
  });

  it('letter allows just once char', () => {
    const actual = command.validate({ options: { letter: "a" } });
    assert.strictEqual(actual, true);
  });

  it('letter allows just once char', () => {
    const actual = command.validate({ options: { letter: "ab" } });
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
});