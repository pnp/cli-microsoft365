import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./task-checklistitem-add');

describe(commands.TASK_CHECKLISTITEM_ADD, () => {
  const validTaskId = 'BC3L9DGJ5UG2UQn4MlEbcZcALpqb';
  const validTitle = 'Checklist item title';

  const taskDetailsResponse = {
    '@odata.etag': 'W/"JzEtVGFza0RldGFpbHMgQEBAQEBAQEBAQEBAQEBBTCc="',
    id: validTaskId
  };

  const taskDetailsWithChecklistResponse: any = {
    id: validTaskId,
    checklist: {}
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(telemetry, 'trackEvent').callsFake(() => { });
    sinon.stub(pid, 'getProcessName').callsFake(() => '');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinonUtil.restore([
      auth.restoreAuth,
      telemetry.trackEvent,
      pid.getProcessName
    ]);
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.TASK_CHECKLISTITEM_ADD), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'isChecked']);
  });

  it('correctly adds checklist item', async () => {
    let checklistItemId = '';

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        checklistItemId = Object.keys(opts.data.checklist)[0];
        taskDetailsWithChecklistResponse.checklist[checklistItemId] = {
          title: validTitle,
          isChecked: false
        };
        return taskDetailsWithChecklistResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        title: validTitle,
        output: 'json'
      }
    });
    assert(loggerLogSpy.calledWith(taskDetailsWithChecklistResponse.checklist[checklistItemId]));
  });

  it('correctly adds checklist item with text output', async () => {
    let checklistItemId = '';
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        checklistItemId = Object.keys(opts.data.checklist)[0];
        taskDetailsWithChecklistResponse.checklist[checklistItemId] = {
          title: validTitle,
          isChecked: false
        };
        return taskDetailsWithChecklistResponse;
      }

      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        title: validTitle,
        output: 'text'
      }
    });
    assert(loggerLogSpy.calledWith({ isChecked: false, title: validTitle, id: checklistItemId }));
  });

  it('fails when unexpected API error was thrown', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        throw 'Something went wrong.';
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        title: validTitle
      }
    }), new CommandError('Something went wrong.'));
  });

  it('fails when Planner task does not exist', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        throw 'The request item is not found.';
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        title: validTitle
      }
    }), new CommandError('Planner task was not found.'));
  });
});
