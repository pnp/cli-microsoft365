import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { Cli } from '../../../../cli/Cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { CommandError } from '../../../../Command.js';
import request from '../../../../request.js';
import { telemetry } from '../../../../telemetry.js';
import { formatting } from '../../../../utils/formatting.js';
import { pid } from '../../../../utils/pid.js';

import { sinonUtil } from '../../../../utils/sinonUtil.js';
import commands from '../../commands.js';
import command from './task-reference-add.js';
import { session } from '../../../../utils/session.js';

describe(commands.TASK_REFERENCE_ADD, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  const validTaskId = '2Vf8JHgsBUiIf-nuvBtv-ZgAAYw2';
  const validUrl = 'https://www.microsoft.com';
  const validAlias = 'Test';
  const validType = 'Word';

  const referenceResponse = {
    "https%3A//www%2Emicrosoft%2Ecom": {
      "alias": "Test",
      "type": "Word",
      "previewPriority": "8585493318091789098Pa",
      "lastModifiedDateTime": "2022-05-11T13:18:56.3142944Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "dd8b99a7-77c6-4238-a609-396d27844921"
        }
      }
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
    commandInfo = Cli.getCommandInfo(command);
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
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.patch
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REFERENCE_ADD);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation if incorrect type is specified.', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl,
        type: "wrong"
      }
    }, commandInfo);
    assert.notStrictEqual(actual, true);
  });

  it('passes validation when valid options specified', async () => {
    const actual = await command.validate({
      options: {
        taskId: validTaskId,
        url: validUrl
      }
    }, commandInfo);
    assert.strictEqual(actual, true);
  });

  it('correctly adds reference', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return { references: referenceResponse };
      }

      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      taskId: validTaskId,
      url: validUrl
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(referenceResponse));
  });

  it('correctly adds reference with type and alias', async () => {
    sinon.stub(request, 'patch').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details`) {
        return { references: referenceResponse };

      }
      throw 'Invalid Request';
    });

    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details` &&
        JSON.stringify(opts.headers) === JSON.stringify({
          'accept': 'application/json'
        })) {
        return { "@odata.etag": "TestEtag" };
      }

      throw 'Invalid Request';
    });

    const options: any = {
      taskId: validTaskId,
      url: validUrl,
      alias: validAlias,
      type: validType
    };

    await command.action(logger, { options: options } as any);
    assert(loggerLogSpy.calledWith(referenceResponse));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
