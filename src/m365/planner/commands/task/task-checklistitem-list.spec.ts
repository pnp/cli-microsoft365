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
import command from './task-checklistitem-list.js';

describe(commands.TASK_CHECKLISTITEM_LIST, () => {
  let log: string[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;

  const jsonOutput = {
    "checklist": {
      "33224": {
        "isChecked": false,
        "title": "Some checklist",
        "orderHint": "8585576049720396756P(",
        "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      },
      "69115": {
        "isChecked": false,
        "title": "Some checklist more",
        "orderHint": "85855760494@",
        "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
        "lastModifiedBy": {
          "user": {
            "displayName": null,
            "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
          }
        }
      }
    }
  };
  const textOutput = {
    "checklist": [{
      "id": "33224",
      "isChecked": false,
      "title": "Some checklist",
      "orderHint": "8585576049720396756P(",
      "lastModifiedDateTime": "2022-02-04T19:12:53.4692149Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    },
    {
      "id": "69115",
      "isChecked": false,
      "title": "Some checklist more",
      "orderHint": "85855760494@",
      "lastModifiedDateTime": "2022-02-04T19:12:55.4735671Z",
      "lastModifiedBy": {
        "user": {
          "displayName": null,
          "id": "88e85b64-e687-4e0b-bbf4-f42f5f8e674e"
        }
      }
    }
    ]
  };

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').returns();
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
      request.get
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
    auth.connection.accessTokens = {};
  });

  it('has correct name', () => {
    assert.strictEqual(command.name, commands.TASK_CHECKLISTITEM_LIST);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('defines correct properties for the default output', () => {
    assert.deepStrictEqual(command.defaultProperties(), ['id', 'title', 'isChecked']);
  });

  it('successfully handles item found(JSON)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true
      }
    });
    assert(loggerLogSpy.calledWith(jsonOutput.checklist));
  });

  it('successfully handles item found(TEXT)', async () => {
    sinon.stub(request, 'get').callsFake(async (opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/planner/tasks/vzCcZoOv-U27PwydxHB8opcADJo-/details?$select=checklist`) {
        return jsonOutput;
      }

      throw 'Invalid request';
    });

    await command.action(logger, {
      options: {
        taskId: 'vzCcZoOv-U27PwydxHB8opcADJo-', debug: true, output: 'text'
      }
    });
    assert(loggerLogSpy.calledWith(textOutput.checklist));
  });

  it('correctly handles item not found', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('The requested item is not found.'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('The requested item is not found.'));
  });

  it('correctly handles random API error', async () => {
    sinonUtil.restore(request.get);
    sinon.stub(request, 'get').rejects(new Error('An error has occurred'));

    await assert.rejects(command.action(logger, { options: {} } as any), new CommandError('An error has occurred'));
  });
});
