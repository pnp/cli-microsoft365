import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './list-add.js';
import { accessToken } from '../../../../utils/accessToken.js';

describe(commands.LIST_ADD, () => {
  let log: string[];
  let logger: Logger;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    sinon.stub(accessToken, 'assertDelegatedAccessToken').returns();
    auth.connection.active = true;
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => {
        log.push(msg);
      },
      logRaw: async (msg: string) => {
        log.push(msg);
      },
      logToStderr: async (msg: string) => {
        log.push(msg);
      }
    };
    (command as any).items = [];
  });

  afterEach(() => {
    sinonUtil.restore([
      request.post,
      Date.now
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.LIST_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('adds To Do task list', async () => {
    sinon.stub(request, 'post').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/me/todo/lists`) {
        return {
          "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists/$entity",
          "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGlTQ==\"",
          "displayName": "FooList",
          "isOwner": true,
          "isShared": false,
          "wellknownListName": "none",
          "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIgAAA="
        };
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        name: "FooList"
      }
    } as any);
    assert.strictEqual(JSON.stringify(log[0]), JSON.stringify({
      "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#lists/$entity",
      "@odata.etag": "W/\"m1fdwWoFiE2YS9yegTKoYwAA/ZGlTQ==\"",
      "displayName": "FooList",
      "isOwner": true,
      "isShared": false,
      "wellknownListName": "none",
      "id": "AAMkAGI3NDhlZmQzLWQxYjAtNGJjNy04NmYwLWQ0M2IzZTNlMDUwNAAuAAAAAACQ1l2jfH6VSZraktP8Z7auAQCbV93BagWITZhL3J6BMqhjAAD9pHIgAAA="
    }));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'post').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { name: "FooList" } } as any), new CommandError('An error has occurred'));
  });
});
