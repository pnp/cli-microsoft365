import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-checklistitem-add.js';

describe(commands.TASK_CHECKLISTITEM_ADD, () => {
  const validTaskId = 'BC3L9DGJ5UG2UQn4MlEbcZcALpqb';
  const validTitle = 'Checklist item title';

  const taskDetailsResponse = {
    '@odata.etag': 'W/"JzEtVGFza0RldGFpbHMgQEBAQEBAQEBAQEBAQEBBTCc="',
    id: validTaskId
  };

  const taskDetailsWithChecklistResponse = {
    id: validTaskId,
    checklist: {
      '00000000-0000-0000-0000-000000000000': {
        title: validTitle,
        isChecked: false
      }
    }
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    auth.connection.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
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
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_CHECKLISTITEM_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('correctly adds checklist item', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
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
    assert(loggerLogSpy.calledWith(taskDetailsWithChecklistResponse.checklist));
  });

  it('correctly adds checklist item with text output', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return taskDetailsResponse;
      }

      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
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
    assert(loggerLogSpy.calledWith([{ id: '00000000-0000-0000-0000-000000000000', title: validTitle, isChecked: false }]));
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
    }), new CommandError('The request item is not found.'));
  });
});
