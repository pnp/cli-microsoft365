import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./list-list');

describe(commands.LIST_LIST, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
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
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['displayName', 'id']);
  });

  it('lists To Do task lists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists",
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
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
      }
    });
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
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));
    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
