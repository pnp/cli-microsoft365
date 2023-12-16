import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-checklistitem-remove.js';

describe(commands.TASK_CHECKLISTITEM_REMOVE, () => {
  let log: string[];
  let logger: Logger;
  let commandInfo: CommandInfo;
  let promptIssued: boolean = false;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validId = '71175';

  const responseChecklistWithId = {
    "71175": {
      "isChecked": false,
      "title": "test 2"
    }
  };
  const responseChecklistWithNoId = {
    "71176": {
      "isChecked": false,
      "title": "test 2"
    }
  };


  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.service.connected = true;
    auth.service.accessTokens[(command as any).resource] = {
      accessToken: 'abc',
      expiresOn: new Date()
    };
    sinon.stub(cli.getConfig(), 'all').value({});
    commandInfo = cli.getCommandInfo(command);
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
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(true);
    });

    promptIssued = false;
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch,
      cli.promptForConfirmation
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_CHECKLISTITEM_REMOVE);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('prompts before removal when force option not passed', async () => {
    sinonUtil.restore(cli.promptForConfirmation);
    sinon.stub(cli, 'promptForConfirmation').callsFake(() => {
      promptIssued = true;
      return Promise.resolve(false);
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    });

    assert(promptIssued);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        id: validId
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly deletes checklist item', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
        return {
          "@odata.etag": "TestEtag",
          checklist: responseChecklistWithId
        };
      }
      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return;
      }
      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId,
        force: true
      }
    });
  });

  it('successfully remove checklist item with confirmation prompt', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
        return {
          "@odata.etag": "TestEtag",
          checklist: responseChecklistWithId
        };
      }
      throw 'Invalid Request';
    });
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return;
      }
      throw 'Invalid Request';
    });

    await command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    });
  });

  it('fails when checklist item does not exists', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=checklist`) {
        return {
          "@odata.etag": "TestEtag",
          checklist: responseChecklistWithNoId
        };
      }

      throw 'Invalid Request';
    });

    await assert.rejects(command.action(logger, {
      options: {
        taskId: validTaskId,
        id: validId
      }
    }), new CommandError(`The specified checklist item with id ${validId} does not exist`));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
