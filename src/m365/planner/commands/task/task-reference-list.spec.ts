import * as assert from 'assert';
import * as sinon from 'sinon';
import { telemetry } from '../../../../telemetry';
import auth from '../../../../Auth';
import { Logger } from '../../../../cli/Logger';
import Command, { CommandError } from '../../../../Command';
import request from '../../../../request';
import { formatting } from '../../../../utils/formatting';
import { pid } from '../../../../utils/pid';
import { session } from '../../../../utils/session';
import { sinonUtil } from '../../../../utils/sinonUtil';
import commands from '../../commands';
const command: Command = require('./task-reference-list');

describe(commands.TASK_REFERENCE_LIST, () => {
  const referenceListResponse = {
    "https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.pdf": {
      "alias": "Sample.pdf",
      "type": "Pdf",
      "previewPriority": "[>",
      "lastModifiedDateTime": "2022-05-15T16:20:31.8649232Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "fe36f75f-c103-410b-a18a-2bf6df06ac3a"
        }
      }
    },
    "https%3A//contoso%2Esharepoint%2Ecom/sites/HRPlan/Shared Documents/Sample.png": {
      "alias": "Sample.png",
      "type": "Other",
      "previewPriority": "8585492445655664725P(",
      "lastModifiedDateTime": "2022-05-12T13:32:59.9267487Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "fe36f75f-c103-410b-a18a-2bf6df06ac3a"
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.service.connected = false;
    auth.service.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_REFERENCE_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('successfully handles item found', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/${formatting.encodeQueryParameter("uBk5fK_MHkeyuPYlCo4OFpcAMowf")}/details?$select=references`) {
        return references;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        taskId: 'uBk5fK_MHkeyuPYlCo4OFpcAMowf'
      }
    });
    assert(loggerLogSpy.calledWith(references.references));
  });

  it('handles error correctly', async () => {
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: { taskId: 'uBk5fK_MHkeyuPYlCo4OFpcAMowf' } } as any), new CommandError('An error has occurred'));
  });
});
