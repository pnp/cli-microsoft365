import * as assert from 'assert';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
import { Cli, CommandInfo, Logger } from '../../../../cli';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { sinonUtil } from '../../../../utils';
import commands from '../../commands';
const command: Command = require('./list-set');

describe(commands.LIST_SET, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    auth.service.connected = true;
    commandInfo = Cli.getCommandInfo(command);
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.LIST_SET), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('updates a To Do list by ID', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists/$entity",
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqngw==\"",
              "displayName": "Bar",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=",
        newName: "Bar"
      }
    } as any, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('updates a To Do list by Name', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'FooList'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGllw==\"",
              "displayName": "FooList",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIiAAA="
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
          "value": [
            {
              "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists/$entity",
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqngw==\"",
              "displayName": "foo",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
            }
          ]
        });
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        name: "FooList",
        newName: "Bar"
      }
    }, () => {
      try {
        assert.strictEqual(log.length, 0);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('Handles error when List name is not available', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists?$filter=displayName eq 'InvalidFooList'`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('8bc5c293-4208-414e-9837-718a0385be6b')/todo/lists",
          "value": []
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'patch').callsFake(() => {
      return Promise.resolve();
    });

    command.action(logger, {
      options: {
        debug: false,
        name: "InvalidFooList",
        newName: "foo"
      }
    }, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('The list InvalidFooList cannot be found')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error correctly', (done) => {
    sinon.stub(request, 'patch').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists/AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=`) {
        return Promise.reject('An error has occurred');
      }

      return Promise.reject('Invalid request');
    });

    command.action(logger, {
      options: {
        debug: false,
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=",
        newName: "Foo"
      }
    } as any, (err?: any) => {
      try {
        assert.strictEqual(JSON.stringify(err), JSON.stringify(new CommandError('An error has occurred')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if new name is not set', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
      }
    }, commandInfo);
    assert.notEqual(actual, true);
  });

  it('fails validation if neither id nor name is not set', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        newName: "Foo"
      }
    }, commandInfo);
    assert.notEqual(actual, true);
  });

  it('fails validation if both id and name are set', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=",
        name: "FooList",
        newName: "Foo"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when all parameters are valid', async () => {
    const actual = await command.validate({
      options: {
        debug: false,
        id: "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA=",
        newName: 'Foo'
      }
    }, commandInfo);

    assert.equal(actual, true);
  });

  it('supports debug mode', () => {
    const options = command.options;
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });
});