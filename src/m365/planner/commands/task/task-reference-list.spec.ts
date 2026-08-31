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
import command, { options } from './task-reference-list.js';

describe(commands.TASK_REFERENCE_LIST, () => {
  const referenceListResponse = {
    'https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.pdf': {
      alias: 'Sample.pdf',
      type: 'Pdf',
      previewPriority: '[>',
      lastModifiedDateTime: '2022-05-15T16:20:31.8649232Z',
      lastModifiedBy: {
        user: {
          displayName: null,
          id: 'fe36f75f-c103-410b-a18a-2bf6df06ac3a'
        }
      }
    },
    'https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.png': {
      alias: 'Sample.png',
      type: 'Other',
      previewPriority: '8585492445655664725P(',
      lastModifiedDateTime: '2022-05-12T13:32:59.9267487Z',
      lastModifiedBy: {
        user: {
          displayName: null,
          id: 'fe36f75f-c103-410b-a18a-2bf6df06ac3a'
        }
      }
    }
  };

  const references = {
    references: [
      referenceListResponse
    ]
  };

  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;
  const validTaskId = 'uBk5fK_MHkeyuPYlCo4OFpcAMowf';

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
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REFERENCE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('fails validation with unknown options', () => {
    const actual = commandOptionsSchema.safeParse({
      taskId: validTaskId,
      unknownOption: 'value'
    });
    assert.strictEqual(actual.success, false);
  });

  it('successfully handles item found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter(validTaskId)}/details?$select=references`) {
        return references;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: commandOptionsSchema.parse({
        taskId: validTaskId
      })
    });
    assert(loggerLogSpy.calledWith(references.references));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, {
      options: commandOptionsSchema.parse({
        taskId: validTaskId
      })
    }), new CommandError('An error has occurred'));
  });
});
