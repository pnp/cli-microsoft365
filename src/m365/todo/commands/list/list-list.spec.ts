import commands from '../../commands';
import Command, { CommandOption } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./list-list');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.LIST_LIST, () => {
  let log: string[];
  let cmdInstance: any;

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
    assert.strictEqual(command.name.startsWith(commands.LIST_LIST), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('lists To Do task lists', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/todo/lists`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
              "displayName": "Tasks",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "defaultList",
              "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
              "displayName": "Foo",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([
          {
            "displayName": "Tasks",
            "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
          },
          {
            "displayName": "Foo",
            "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
          },
          {
            "displayName": "Bar",
            "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
          }
        ]);
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('shows all lists data in JSON output mode', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/todo/lists`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
              "displayName": "Tasks",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "defaultList",
              "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
              "displayName": "Foo",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        output: 'json'
      }
    }, () => {
      try {
        const actual = JSON.stringify(log[log.length - 1]);
        const expected = JSON.stringify([
          {
            "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
            "displayName": "Tasks",
            "isOwner": true,
            "isShared": false,
            "wellknownListName": "defaultList",
            "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
          },
          {
            "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
            "displayName": "Foo",
            "isOwner": true,
            "isShared": false,
            "wellknownListName": "none",
            "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
          },
          {
            "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
            "displayName": "Bar",
            "isOwner": true,
            "isShared": false,
            "wellknownListName": "none",
            "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIjAAA="
          }
        ]);
        assert.strictEqual(actual, expected);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('lists To Do task lists (verbose)', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/beta/me/todo/lists`) {
        return Promise.resolve({
          "@odata.context": "https://graph.microsoft.com/beta/$metadata#lists",
          "value": [
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpQ==\"",
              "displayName": "Tasks",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "defaultList",
              "id": "AQMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAADkNZdo3x_lUma2pLT-Ge2rgEAm1fdwWoFiE2YS9yegTKoYwAAAgESAAAA"
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrpw==\"",
              "displayName": "Foo",
              "isOwner": true,
              "isShared": false,
              "wellknownListName": "none",
              "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIeAAA="
            },
            {
              "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/hqrqQ==\"",
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

    cmdInstance.action = command.action();
    cmdInstance.action({
      options: {
        debug: false,
        verbose: true
      }
    }, () => {
      try {
        assert(log[log.length - 1].indexOf('DONE') > -1);
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
});